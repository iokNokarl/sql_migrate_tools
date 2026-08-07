"""
数据迁移核心模块
负责将源数据库数据迁移到目标数据库，支持：
- 单表/多表/多库迁移
- 分批处理（超过1000条每批500条）
- 多线程处理（超过5万条时启用，最多20个线程）
- 事务保障，异常时回滚
- PostgreSQL 批量写入优化（execute_values，性能提升 10-30 倍）
- 键集分页（Keyset Pagination，支持字符串主键，O(1) 每批）
"""

import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Tuple
import sys
from .database_connector import DatabaseConnector, create_connector
from .sql_dialect import SQLDialect
from .progress_bar import ProgressBar


BATCH_THRESHOLD = 3000        # 超过此条数启用分批
BATCH_SIZE = 10000             # 每批处理条数
ASYNC_THRESHOLD = 50000       # 超过此条数启用异步线程
MAX_THREADS = 20              # 最大线程数


def _format_pk_value(value, is_integer: bool) -> str:
    """格式化主键值用于 SQL 拼接（键集分页用），正确处理转义"""
    if is_integer:
        return str(int(value))
    s = str(value).replace("'", "''")
    return f"'{s}'"


def _quote_identifier(name: str, db_type: str) -> str:
    """根据数据库类型给标识符加引号"""
    db = db_type.lower()
    if db in ("mysql",):
        return f"`{name}`"
    elif db in ("postgresql", "postgres", "pg", "psql"):
        return f'"{name}"'
    return name


class MigrationError(Exception):
    """迁移异常类"""

    def __init__(self, message: str, table: Optional[str] = None, database: Optional[str] = None):
        self.table = table
        self.database = database
        super().__init__(message)


class MigrationLogger:
    """简单的日志打印器"""

    _lock = threading.Lock()

    @classmethod
    def info(cls, msg: str) -> None:
        with cls._lock:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            print(f"[INFO] {timestamp} {msg}")

    @classmethod
    def warn(cls, msg: str) -> None:
        with cls._lock:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            print(f"[WARN] {timestamp} {msg}")

    @classmethod
    def error(cls, msg: str) -> None:
        with cls._lock:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            print(f"[ERROR] {timestamp} {msg}")

    @classmethod
    def success(cls, msg: str) -> None:
        with cls._lock:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            print(f"[SUCCESS] {timestamp} {msg}")


class DatabaseMigrator:
    """数据库迁移器"""

    def __init__(
        self,
        source_connector: DatabaseConnector,
        target_connector: DatabaseConnector,
        max_threads: int = MAX_THREADS,
        batch_size: int = BATCH_SIZE,
    ):
        self.source = source_connector
        self.target = target_connector
        self.max_threads = min(max_threads, MAX_THREADS)
        self.batch_size = batch_size
        self._migration_lock = threading.Lock()
        self._error_occurred = False
        self._error_message: Optional[str] = None

    def _collect_migration_plan(
        self,
        databases: Optional[List[str]] = None,
        tables: Optional[List[str]] = None,
    ) -> List[Tuple[Optional[str], str]]:
        """
        收集需要迁移的表列表
        返回: [(database_name, table_name), ...]
        """
        plan: List[Tuple[Optional[str], str]] = []

        if tables:
            for tbl in tables:
                plan.append((databases[0] if databases else None, tbl))
            return plan

        if databases:
            for db in databases:
                try:
                    db_tables = self.source.get_tables(db)
                    for tbl in db_tables:
                        plan.append((db, tbl))
                except Exception as exc:
                    MigrationLogger.warn(f"读取数据库 {db} 表清单失败: {exc}")
            return plan

        default_tables = self.source.get_tables()
        for tbl in default_tables:
            plan.append((None, tbl))
        return plan

    def _check_and_create_table(
        self,
        source_db: Optional[str],
        source_table: str,
        target_table: str,
    ) -> List[str]:
        """在目标库检查并创建表结构，返回列名列表
        source_table: 源数据库中的表名（读数据用）
        target_table: 目标数据库中的表名（写数据用）
        """
        schema = self.source.get_table_schema(source_table)
        if not schema:
            raise MigrationError(f"无法获取表 {source_table} 的结构", table=source_table)

        create_sql = SQLDialect.build_create_table_sql(
            table_name=target_table,
            schema=schema,
            target_db=self.target.db_type,
        )
        drop_sql = SQLDialect.build_drop_table_sql(target_table, self.target.db_type)

        try:
            self.target.execute(drop_sql)
            self.target.execute(create_sql)
            if source_table == target_table:
                MigrationLogger.info(f"表 {target_table} 已在目标库创建/重建")
            else:
                MigrationLogger.info(f"表 {source_table} → {target_table} 已在目标库创建/重建")
        except Exception as exc:
            raise MigrationError(
                f"创建表 {target_table} 失败: {exc}", table=target_table
            ) from exc

        return [col[0] for col in schema]

    def _migrate_single_table(
        self,
        source_db: Optional[str],
        source_table: str,
        target_table: str,
    ) -> Dict[str, Any]:
        """迁移单张表
        source_table: 源数据库中的表名
        target_table: 目标数据库中的表名
        """
        if self._error_occurred:
            raise MigrationError("检测到其他表迁移失败，当前表已中止")

        result = {
            "database": source_db,
            "table": source_table,
            "target_table": target_table,
            "row_count": 0,
            "batch_count": 0,
            "thread_count": 1,
            "elapsed": 0.0,
        }
        start_time = time.time()

        row_count = self.source.get_row_count(source_table)
        result["row_count"] = row_count
        if source_table == target_table:
            MigrationLogger.info(f"开始迁移表 {source_table}，共 {row_count} 条记录")
        else:
            MigrationLogger.info(
                f"开始迁移表 {source_table} → {target_table}，共 {row_count} 条记录"
            )

        if row_count == 0:
            self._check_and_create_table(source_db, source_table, target_table)
            result["elapsed"] = time.time() - start_time
            if source_table == target_table:
                MigrationLogger.success(f"表 {target_table} 迁移完成（空表）")
            else:
                MigrationLogger.success(f"表 {source_table} → {target_table} 迁移完成（空表）")
            return result

        columns = self._check_and_create_table(source_db, source_table, target_table)
        use_batched = row_count > BATCH_THRESHOLD
        # SQLite 是文件型数据库，多线程会因文件锁竞争变慢甚至卡死 → 始终单线程
        source_is_sqlite = getattr(self.source, "db_type", "").lower() == "sqlite"
        use_async = (row_count > ASYNC_THRESHOLD) and not source_is_sqlite

        if use_batched:
            batch_count = (row_count + self.batch_size - 1) // self.batch_size
            result["batch_count"] = batch_count
            MigrationLogger.info(
                f"数据量 {row_count} 超过 {BATCH_THRESHOLD} 条，"
                f"采用分批处理（每批 {self.batch_size} 条，共 {batch_count} 批）"
            )

            if use_async:
                thread_count = min(self.max_threads, max(1, batch_count))
                result["thread_count"] = thread_count
                MigrationLogger.info(
                    f"数据量 {row_count} 超过 {ASYNC_THRESHOLD} 条，"
                    f"启用 {thread_count} 个并发线程处理"
                )
                self._migrate_batched_async(source_table, target_table, columns, row_count, batch_count)
            else:
                if source_is_sqlite and row_count > ASYNC_THRESHOLD:
                    MigrationLogger.info(
                        f"源库为 SQLite（文件型数据库），多线程会因文件锁竞争变慢，"
                        f"已使用高效单线程分批模式"
                    )
                self._migrate_batched_sync(source_table, target_table, columns, row_count, batch_count)
        else:
            self._migrate_whole_table(source_table, target_table, columns)
        result["elapsed"] = time.time() - start_time
        if source_table == target_table:
            MigrationLogger.success(
                f"表 {target_table} 迁移完成，耗时 {result['elapsed']:.2f}s，"
                f"共处理 {row_count} 条记录"
            )
        else:
            MigrationLogger.success(
                f"表 {source_table} → {target_table} 迁移完成，耗时 {result['elapsed']:.2f}s，"
                f"共处理 {row_count} 条记录"
            )
        return result

    def _migrate_whole_table(self, source_table: str, target_table: str,
                             columns: List[str]) -> None:
        """整表一次性迁移（仅用于小数据量）"""
        select_sql = SQLDialect.build_select_batched_sql(
            table_name=source_table,
            columns=columns,
            batch_size=BATCH_SIZE,
            offset=0,
            db_type=self.source.db_type,
        )
        select_sql = select_sql.split(" LIMIT")[0] + ";"
        rows = self.source.fetchall(select_sql)
        if not rows:
            return
        # 小量数据也给个快速进度条（以行数为单位）
        total = len(rows)
        display = target_table if (source_table == target_table) else f"{source_table}→{target_table}"
        with ProgressBar(total=total, prefix=f"  {display}", unit="条",
                          min_update_interval=0.2):
            self._insert_batch(target_table, columns, list(rows))

    def _detect_primary_key(self, table: str) -> Tuple[Optional[str], bool]:
        """
        检测表是否有"单列主键"，返回 (主键列名, 是否整数类型)。

        有整数主键 → 可以用主键范围分页（O(1)，超快速，270 万行 < 2 分钟）
        有字符串主键 → 可以用键集分页（O(1) 每批，比 LIMIT/OFFSET 快 10-50 倍）
        无主键 → 只能用 LIMIT/OFFSET（O(n²)，大表超慢）
        """
        try:
            schema = self.source.get_table_schema(table)
        except Exception:
            return None, False

        pk_cols = [col for col, col_type, is_nullable, is_primary in schema if is_primary]
        if len(pk_cols) != 1:
            return None, False

        pk_col = pk_cols[0]
        pk_type_col = None
        for col, col_type, is_nullable, is_primary in schema:
            if is_primary:
                pk_type_col = (col_type or "").upper()
                break

        if pk_type_col is None:
            return None, False

        integer_keywords = ("INT", "INTEGER", "BIGINT", "SMALLINT", "TINYINT", "MEDIUMINT", "SERIAL")
        is_integer = any(k in pk_type_col for k in integer_keywords)

        return pk_col, is_integer

    def _detect_integer_primary_key(self, table: str) -> Optional[str]:
        """兼容旧接口：返回整数主键列名，非整数返回 None"""
        pk_col, is_integer = self._detect_primary_key(table)
        return pk_col if is_integer else None

    def _split_pk_ranges(self, min_val: int, max_val: int, batch_size: int) -> List[Tuple[int, int]]:
        """根据主键范围 + 批大小，分割成 [start, end) 区间列表。"""
        if min_val > max_val:
            return []
        ranges: List[Tuple[int, int]] = []
        start = min_val
        while start <= max_val:
            end = min(start + batch_size, max_val + 1)
            ranges.append((start, end))
            start = end
        return ranges

    def _migrate_batched_sync(
        self,
        source_table: str,
        target_table: str,
        columns: List[str],
        total_rows: int,
        batch_count: int,
    ) -> None:
        """同步分批迁移（单事务，全部成功才 commit）
        - 有整数主键 → 主键范围查询（O(1)，最快）
        - 有字符串主键 → 键集分页（O(1) 每批，快）
        - 无主键 → LIMIT/OFFSET（O(n²)，大表慢）
        """
        display = target_table if (source_table == target_table) else f"{source_table}→{target_table}"
        bar = ProgressBar(total=total_rows, prefix=f"  {display}",
                          unit="条", min_update_interval=0.15)
        try:
            # 先检测主键类型（性能关键）
            pk_col, pk_is_integer = self._detect_primary_key(source_table)
            pk_ranges = None
            use_keyset = False

            if pk_col is not None and pk_is_integer:
                # --- 整数主键：范围查询 ---
                try:
                    row = self.source.fetchone(
                        f"SELECT MIN({pk_col}), MAX({pk_col}) FROM {source_table}"
                    )
                    if row is not None and row[0] is not None:
                        pk_min, pk_max = row[0], row[1]
                        pk_ranges = self._split_pk_ranges(pk_min, pk_max, self.batch_size)
                        print(
                            f"[INFO] {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} "
                            f"  检测到整数主键 {pk_col}，"
                            f"使用主键范围查询（O(1)，超快速，"
                            f"共 {len(pk_ranges)} 段）",
                            flush=True,
                        )
                except Exception:
                    pass
            elif pk_col is not None:
                # --- 字符串主键：键集分页 ---
                use_keyset = True
                print(
                    f"[INFO] {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} "
                    f"  检测到字符串主键 {pk_col}，"
                    f"使用键集分页（O(1) 每批，比 LIMIT/OFFSET 快 10-50 倍）",
                    flush=True,
                )

            # 在干净的新行上渲染进度条，避免 \r 在 PowerShell 中的兼容问题
            bar._render(force=True)

            if pk_ranges is not None:
                # --- 整数主键范围查询：O(1) ---
                quoted_columns = [SQLDialect.quote_identifier(c, self.source.db_type) for c in columns]
                col_sql = ", ".join(quoted_columns)
                quoted_pk = SQLDialect.quote_identifier(pk_col, self.source.db_type)
                for start_val, end_val in pk_ranges:
                    if self._error_occurred:
                        raise MigrationError("检测到错误，停止分批迁移")
                    select_sql = (
                        f"SELECT {col_sql} FROM {SQLDialect.quote_identifier(source_table, self.source.db_type)} "
                        f"WHERE {quoted_pk} >= {int(start_val)} AND {quoted_pk} < {int(end_val)} "
                        f"ORDER BY {quoted_pk}"
                    )
                    rows = self.source.fetchall(select_sql)
                    if rows:
                        self._insert_batch(target_table, columns, list(rows))
                        bar.add(len(rows))
            elif use_keyset:
                # --- 键集分页：O(1) 每批 ---
                q_pk = _quote_identifier(pk_col, self.source.db_type)
                q_table = SQLDialect.quote_identifier(source_table, self.source.db_type)
                quoted_columns = [SQLDialect.quote_identifier(c, self.source.db_type) for c in columns]
                col_sql = ", ".join(quoted_columns)
                last_pk_val = None
                try:
                    pk_idx = columns.index(pk_col)
                except ValueError:
                    pk_idx = 0
                written = 0
                while written < total_rows:
                    if self._error_occurred:
                        raise MigrationError("检测到错误，停止分批迁移")
                    if last_pk_val is None:
                        select_sql = (f"SELECT {col_sql} FROM {q_table} "
                                      f"ORDER BY {q_pk} LIMIT {self.batch_size}")
                    else:
                        formatted = _format_pk_value(last_pk_val, pk_is_integer)
                        select_sql = (f"SELECT {col_sql} FROM {q_table} "
                                      f"WHERE {q_pk} > {formatted} "
                                      f"ORDER BY {q_pk} LIMIT {self.batch_size}")
                    rows = self.source.fetchall(select_sql)
                    if not rows:
                        break
                    batch = list(rows)
                    written += len(batch)
                    last_pk_val = batch[-1][pk_idx]
                    self._insert_batch(target_table, columns, batch)
                    bar.add(len(batch))
            else:
                # --- 回退：LIMIT/OFFSET ---
                print(
                    f"[WARN] {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} "
                    f"  未检测到主键，使用 LIMIT/OFFSET（大表可能较慢，"
                    f"建议添加主键列以提升迁移速度）"
                )
                for batch_idx in range(batch_count):
                    if self._error_occurred:
                        raise MigrationError("检测到错误，停止分批迁移")
                    offset = batch_idx * self.batch_size
                    limit = min(self.batch_size, total_rows - offset)
                    select_sql = SQLDialect.build_select_batched_sql(
                        table_name=source_table,
                        columns=columns,
                        batch_size=limit,
                        offset=offset,
                        db_type=self.source.db_type,
                    )
                    rows = self.source.fetchall(select_sql)
                    if rows:
                        self._insert_batch(target_table, columns, list(rows))
                    bar.add(len(rows) if rows else limit)
        finally:
            bar.finish()

    def _migrate_batched_async(
        self,
        source_table: str,
        target_table: str,
        columns: List[str],
        total_rows: int,
        batch_count: int,
    ) -> None:
        """多线程分批迁移
        - 多线程只读：每个线程独立建源库连接，并发读取不同分区
        - 写入串行：主线程通过共享目标连接写入（PostgreSQL 使用 execute_values 优化）
        - 乱序安全：收到一批立即写入并更新进度条，不等前批到达
        - 全局事务：全部批次成功后才 commit，中途异常或中断则 rollback
        - 优先使用主键范围查询（O(1)），字符串主键用键集分页（O(1)），无主键则回退 LIMIT/OFFSET
        """
        from queue import Queue

        # 先检测主键类型（性能关键）
        pk_col, pk_is_integer = self._detect_primary_key(source_table)
        use_pk_range = False
        use_keyset = False
        total_tasks = batch_count

        if pk_col is not None and pk_is_integer:
            try:
                row = self.source.fetchone(
                    f"SELECT MIN({pk_col}), MAX({pk_col}) FROM {source_table}"
                )
                if row is not None and row[0] is not None:
                    pk_min, pk_max = row[0], row[1]
                    pk_ranges = self._split_pk_ranges(pk_min, pk_max, self.batch_size)
                    use_pk_range = True
                    total_tasks = len(pk_ranges)
            except Exception:
                pass
        elif pk_col is not None:
            use_keyset = True

        # 工作队列（无界队列，避免生产者阻塞）
        #   ("offset", batch_idx, offset, limit)        → LIMIT/OFFSET
        #   ("range",  batch_idx, start_val, end_val)   → 主键范围
        work_queue: "Queue[Any]" = Queue()
        result_queue: "Queue[Any]" = Queue()

        write_lock = threading.Lock()
        abort_event = threading.Event()

        def worker_thread(worker_id: int) -> None:
            """工作线程：只负责从源库读取数据，不做写入"""
            try:
                local_source = create_connector(
                    self.source.db_type, dict(getattr(self.source, "config", {}))
                )
                local_source.connect()
            except Exception as exc:
                self._set_error(f"工作线程 {worker_id} 无法连接源数据库: {exc}")
                abort_event.set()
                while not work_queue.empty():
                    try:
                        work_queue.get_nowait()
                    except Exception:
                        pass
                result_queue.put(("__error__", str(exc)))
                return

            try:
                quoted_columns = [SQLDialect.quote_identifier(c, local_source.db_type) for c in columns]
                col_sql = ", ".join(quoted_columns)
                quoted_table = SQLDialect.quote_identifier(source_table, local_source.db_type)
                if use_pk_range:
                    quoted_pk = SQLDialect.quote_identifier(pk_col, local_source.db_type)
                if use_keyset:
                    q_pk = _quote_identifier(pk_col, local_source.db_type)
                    try:
                        keyset_pk_idx = columns.index(pk_col)
                    except ValueError:
                        keyset_pk_idx = 0

                while not abort_event.is_set():
                    try:
                        item = work_queue.get(timeout=1.0)
                    except Exception:
                        return

                    task_type = item[0]
                    batch_idx = item[1]

                    try:
                        if task_type == "range":
                            start_val, end_val = item[2], item[3]
                            select_sql = (
                                f"SELECT {col_sql} FROM {quoted_table} "
                                f"WHERE {quoted_pk} >= {int(start_val)} AND {quoted_pk} < {int(end_val)} "
                                f"ORDER BY {quoted_pk}"
                            )
                            rows = local_source.fetchall(select_sql)
                            result_queue.put((batch_idx, list(rows)))
                        elif task_type == "keyset":
                            # 键集分页：顺序读取，每批放入队列
                            last_pk_val = None
                            written = 0
                            keyset_batch_idx = 0
                            while written < total_rows and not abort_event.is_set():
                                if last_pk_val is None:
                                    select_sql = (f"SELECT {col_sql} FROM {quoted_table} "
                                                  f"ORDER BY {q_pk} LIMIT {self.batch_size}")
                                else:
                                    formatted = _format_pk_value(last_pk_val, pk_is_integer)
                                    select_sql = (f"SELECT {col_sql} FROM {quoted_table} "
                                                  f"WHERE {q_pk} > {formatted} "
                                                  f"ORDER BY {q_pk} LIMIT {self.batch_size}")
                                rows = local_source.fetchall(select_sql)
                                if not rows:
                                    break
                                batch = list(rows)
                                written += len(batch)
                                last_pk_val = batch[-1][keyset_pk_idx]
                                result_queue.put((keyset_batch_idx, batch))
                                keyset_batch_idx += 1
                            # 发送 sentinel 标记键集分页完成
                            result_queue.put((-1, None))
                        else:
                            offset, limit = item[2], item[3]
                            select_sql = SQLDialect.build_select_batched_sql(
                                table_name=source_table,
                                columns=columns,
                                batch_size=limit,
                                offset=offset,
                                db_type=local_source.db_type,
                            )
                            rows = local_source.fetchall(select_sql)
                            result_queue.put((batch_idx, list(rows)))
                    except Exception as exc:
                        self._set_error(f"批次 {batch_idx} 读取失败: {exc}")
                        abort_event.set()
                        result_queue.put(("__error__", f"批次 {batch_idx}: {exc}"))
                        return
            finally:
                try:
                    local_source.close()
                except Exception:
                    pass

        # ---- 先创建进度条，打印日志，再启动工作线程 ----
        display = target_table if (source_table == target_table) else f"{source_table}→{target_table}"
        bar = ProgressBar(total=total_rows, prefix=f"  {display}",
                          unit="条", min_update_interval=0.1)

        actual_threads = min(self.max_threads, max(1, total_tasks))

        # 先填充工作队列（无界队列，不会阻塞），再启动工作线程
        # 这样工作线程启动后能立即取到任务，不会因队列空而提前退出
        if use_pk_range:
            for batch_idx, (start_val, end_val) in enumerate(pk_ranges):
                work_queue.put(("range", batch_idx, start_val, end_val))
        elif use_keyset:
            total_tasks = 1
            work_queue.put(("keyset", 0, pk_col, pk_is_integer))
        else:
            MigrationLogger.warn(
                f"  未检测到主键，使用 LIMIT/OFFSET（大表可能较慢，"
                f"建议添加主键列以提升迁移速度）"
            )
            for batch_idx in range(batch_count):
                offset = batch_idx * self.batch_size
                limit = min(self.batch_size, total_rows - offset)
                work_queue.put(("offset", batch_idx, offset, limit))

        if use_keyset:
            print(
                f"[INFO] {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} "
                f"表 {source_table} → {target_table} 使用键集分页读取，"
                f"写入在全局事务中统一提交",
                flush=True,
            )
        else:
            print(
                f"[INFO] {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} "
                f"表 {source_table} → {target_table} 启动 {actual_threads} 个读取线程，"
                f"写入在全局事务中统一提交",
                flush=True,
            )

        # 在干净的新行上渲染进度条
        bar._render(force=True)
        sys.stdout.flush()
        # 启动工作线程（队列已预填充，工作线程启动后能立即取到任务）
        threads = []
        for i in range(actual_threads):
            t = threading.Thread(target=worker_thread, args=(i,), daemon=True)
            t.start()
            threads.append(t)

        try:
            if use_keyset:
                # 键集分页模式：持续读取直到 sentinel
                received = 0
                while True:
                    if abort_event.is_set() or self._error_occurred:
                        raise MigrationError(
                            f"表 {source_table} → {target_table} 迁移被中止: "
                            f"{self._error_message or '用户中断'}",
                            table=target_table,
                        )
                    try:
                        item = result_queue.get(timeout=0.5)
                    except Exception:
                        if not any(t.is_alive() for t in threads):
                            break
                        bar._render(force=True)  # 刷新进度条时间，让用户知道程序还在运行
                        continue

                    if item[0] == "__error__":
                        abort_event.set()
                        raise MigrationError(
                            f"表 {source_table} → {target_table} 迁移失败: {item[1]}",
                            table=target_table,
                        )

                    batch_idx, rows = item
                    if batch_idx == -1 and rows is None:
                        break  # sentinel: 键集分页完成

                    if rows:
                        with write_lock:
                            self._insert_batch(target_table, columns, rows)
                        bar.add(len(rows))
                        received += 1
            else:
                for received in range(total_tasks):
                    while True:
                        if abort_event.is_set() or self._error_occurred:
                            raise MigrationError(
                                f"表 {source_table} → {target_table} 迁移被中止: "
                                f"{self._error_message or '用户中断'}",
                                table=target_table,
                            )
                        try:
                            item = result_queue.get(timeout=0.2)
                        except Exception:
                            any_alive = any(t.is_alive() for t in threads)
                            if not any_alive:
                                raise MigrationError(
                                    f"表 {source_table} → {target_table} 所有读取线程退出但数据不完整",
                                    table=target_table,
                                )
                            bar._render(force=True)  # 刷新进度条时间，让用户知道程序还在运行
                            continue
                        break

                    if item[0] == "__error__":
                        abort_event.set()
                        raise MigrationError(
                            f"表 {source_table} → {target_table} 迁移失败: {item[1]}",
                            table=target_table,
                        )

                    batch_idx, rows = item
                    if rows:
                        with write_lock:
                            self._insert_batch(target_table, columns, rows)
                        bar.add(len(rows))

            if self._error_occurred or abort_event.is_set():
                raise MigrationError(
                    f"表 {source_table} → {target_table} 异步迁移被中止: "
                    f"{self._error_message or '用户中断'}",
                    table=target_table,
                )

            MigrationLogger.info(
                f"表 {source_table} → {target_table} 全部批次已写入，等待全局事务提交"
            )

        except (Exception, KeyboardInterrupt) as exc:
            abort_event.set()
            for t in threads:
                t.join(timeout=2)
            bar.finish()
            if isinstance(exc, KeyboardInterrupt):
                raise
            raise MigrationError(
                f"表 {source_table} → {target_table} 异步迁移失败: {exc}",
                table=target_table,
            ) from exc

        bar.finish()

        for t in threads:
            t.join(timeout=5)

    def _insert_batch(
        self,
        table_name: str,
        columns: List[str],
        rows: List[tuple],
        target_connector: Optional[DatabaseConnector] = None,
    ) -> None:
        """批量插入数据，PostgreSQL 使用 execute_values 优化（性能提升 10-30 倍）"""
        if not rows:
            return
        conn = target_connector or self.target
        if conn.db_type in ("postgresql", "postgres", "pg", "psql"):
            self._insert_batch_postgresql(conn, table_name, columns, rows)
        else:
            insert_sql = SQLDialect.build_insert_sql(table_name, columns, conn.db_type)
            try:
                conn.executemany(insert_sql, rows)
            except Exception as exc:
                raise MigrationError(
                    f"插入数据失败（表 {table_name}，{len(rows)} 行）: {exc}",
                    table=table_name,
                ) from exc

    def _insert_batch_postgresql(
        self,
        conn: DatabaseConnector,
        table_name: str,
        columns: List[str],
        rows: List[tuple],
    ) -> None:
        """PostgreSQL 高性能批量插入，使用 execute_values（单次网络往返）
        execute_values 要求 SQL 中 VALUES 子句用 %s 占位符，而非 $1/$2
        """
        from psycopg2.extras import execute_values

        quoted_table = SQLDialect.quote_identifier(table_name, conn.db_type)
        quoted_columns = [SQLDialect.quote_identifier(c, conn.db_type) for c in columns]
        insert_sql = (
            f"INSERT INTO {quoted_table} ({', '.join(quoted_columns)}) VALUES %s"
        )
        cursor = conn.cursor()
        try:
            execute_values(cursor, insert_sql, rows, page_size=1000)
        except Exception as exc:
            raise MigrationError(
                f"PostgreSQL 批量插入失败（表 {table_name}，{len(rows)} 行）: {exc}",
                table=table_name,
            ) from exc

    def _insert_batch_with_connector(
        self,
        target_connector: DatabaseConnector,
        table_name: str,
        columns: List[str],
        rows: List[tuple],
    ) -> None:
        """使用指定连接器插入数据（委托给 _insert_batch 统一处理）"""
        self._insert_batch(table_name, columns, rows, target_connector=target_connector)

    def _set_error(self, msg: str) -> None:
        with self._migration_lock:
            if not self._error_occurred:
                self._error_occurred = True
                self._error_message = msg
                MigrationLogger.error(msg)

    def migrate(
        self,
        databases: Optional[List[str]] = None,
        tables: Optional[List[str]] = None,
        table_rename: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """
        执行整体迁移
        所有操作在事务中完成，任何一张表失败（或被 Ctrl+C 中断）
        则全部回滚——目标数据库保持原始状态不变。

        参数：
        - table_rename: 表名映射字典（源表名 → 目标表名），如 {"users": "t_users"}
        """
        overall_result = {
            "success": False,
            "tables": [],
            "total_rows": 0,
            "elapsed": 0.0,
            "error": None,
        }
        overall_start = time.time()

        try:
            self.source.connect()
            self.target.connect()
            MigrationLogger.info(
                f"源数据库: {self.source.db_type} -> 目标数据库: {self.target.db_type}"
            )
            MigrationLogger.info(
                "事务模式: 全局事务（全部成功才提交，中断/异常则全部回滚）"
            )

            plan = self._collect_migration_plan(databases, tables)
            if not plan:
                raise MigrationError("未找到任何需要迁移的表")
            MigrationLogger.info(f"迁移计划共 {len(plan)} 张表")

            for source_db, table_name in plan:
                if self._error_occurred:
                    raise MigrationError(
                        f"迁移过程中发生错误: {self._error_message}"
                    )
                # 计算目标表名：如果 table_rename 中存在映射，则用目标表名
                target_table = table_rename.get(table_name, table_name) if table_rename else table_name

                try:
                    table_result = self._migrate_single_table(source_db, table_name, target_table)
                    overall_result["tables"].append(table_result)
                    overall_result["total_rows"] += table_result["row_count"]
                except Exception as exc:
                    self._set_error(str(exc))
                    raise MigrationError(
                        f"表 {table_name} 迁移失败: {exc}", table=table_name
                    ) from exc

            self.target.commit()
            overall_result["success"] = True
            overall_result["elapsed"] = time.time() - overall_start
            MigrationLogger.success(
                f"全部迁移完成！共处理 {len(overall_result['tables'])} 张表，"
                f"{overall_result['total_rows']} 条记录，"
                f"总耗时 {overall_result['elapsed']:.2f}s"
            )
            return overall_result

        except (Exception, KeyboardInterrupt) as exc:
            overall_result["error"] = str(exc)
            if isinstance(exc, KeyboardInterrupt):
                MigrationLogger.error("检测到用户中断 (Ctrl+C)，正在回滚全部变更...")
            else:
                MigrationLogger.error(f"迁移失败: {exc}，正在回滚全部变更...")
            # 设置中止标记，通知多线程立即退出
            self._error_occurred = True
            try:
                self.target.rollback()
                MigrationLogger.warn("目标数据库已成功回滚，所有变更已放弃")
            except Exception as rollback_exc:
                MigrationLogger.error(f"回滚失败: {rollback_exc}")
            # 如果是用户中断，重新抛出让上层脚本知道是中断
            if isinstance(exc, KeyboardInterrupt):
                raise
            raise

        finally:
            try:
                self.source.close()
            except Exception:
                pass
            try:
                self.target.close()
            except Exception:
                pass
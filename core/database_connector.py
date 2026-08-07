"""
数据库连接管理器模块
负责管理 SQLite / MySQL / PostgreSQL 三种数据库的连接与操作
"""

import os
import sqlite3
from typing import Any, Dict, List, Optional, Tuple


class DatabaseConnector:
    """数据库连接器基类"""

    def __init__(self, db_type: str, config: Dict[str, Any]):
        self.db_type = db_type.lower()
        self.config = config
        self.connection = None
        self._validate_db_type()

    def _validate_db_type(self) -> None:
        """校验数据库类型是否受支持"""
        supported = {"sqlite", "mysql", "postgresql", "psql", "pg"}
        if self.db_type not in supported:
            raise ValueError(
                f"不支持的数据库类型: {self.db_type}，仅支持 sqlite/mysql/postgresql"
            )

    def connect(self) -> Any:
        """建立数据库连接"""
        raise NotImplementedError("子类必须实现 connect 方法")

    def close(self) -> None:
        """关闭数据库连接"""
        if self.connection is not None:
            try:
                self.connection.close()
            except Exception:
                pass
            self.connection = None

    def __enter__(self) -> "DatabaseConnector":
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def cursor(self) -> Any:
        """获取游标"""
        if self.connection is None:
            self.connect()
        return self.connection.cursor()

    def commit(self) -> None:
        """提交事务"""
        if self.connection is not None:
            self.connection.commit()

    def rollback(self) -> None:
        """回滚事务"""
        if self.connection is not None:
            try:
                self.connection.rollback()
            except Exception:
                pass

    def execute(self, sql: str, params: Optional[tuple] = None) -> Any:
        """执行 SQL 语句"""
        cursor = self.cursor()
        if params:
            cursor.execute(sql, params)
        else:
            cursor.execute(sql)
        return cursor

    def executemany(self, sql: str, params: List[tuple]) -> None:
        """批量执行 SQL 语句"""
        cursor = self.cursor()
        cursor.executemany(sql, params)

    def fetchall(self, sql: str, params: Optional[tuple] = None) -> List[tuple]:
        """查询全部数据（统一返回元组列表，兼容 DictCursor 和 TupleCursor）"""
        cursor = self.execute(sql, params)
        rows = cursor.fetchall()
        if rows and isinstance(rows[0], dict):
            # DictCursor 返回字典列表，统一转为元组列表（保持列顺序一致）
            return [tuple(row.values()) for row in rows]
        return rows

    def fetchone(self, sql: str, params: Optional[tuple] = None) -> Optional[tuple]:
        """查询单条数据（统一返回元组，兼容 DictCursor 和 TupleCursor）"""
        cursor = self.execute(sql, params)
        row = cursor.fetchone()
        if row is None:
            return None
        if isinstance(row, dict):
            return tuple(row.values())
        return row


class SQLiteConnector(DatabaseConnector):
    """SQLite 数据库连接器"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__("sqlite", config)
        self.db_path = config.get("path") or config.get("db_path") or config.get("database")
        if not self.db_path:
            raise ValueError("SQLite 需要指定数据库文件路径 (path)")

    def connect(self) -> sqlite3.Connection:
        if not os.path.exists(os.path.dirname(os.path.abspath(self.db_path))):
            parent_dir = os.path.dirname(os.path.abspath(self.db_path))
            if parent_dir and not os.path.exists(parent_dir):
                os.makedirs(parent_dir, exist_ok=True)
        self.connection = sqlite3.connect(
            self.db_path,
            check_same_thread=False,
            timeout=30,
        )
        self.connection.execute("PRAGMA foreign_keys = OFF;")
        return self.connection

    def get_databases(self) -> List[str]:
        return [os.path.basename(self.db_path)]

    def get_tables(self, database: Optional[str] = None) -> List[str]:
        rows = self.fetchall(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;"
        )
        return [row[0] for row in rows if not row[0].startswith("sqlite_")]

    def get_table_schema(self, table_name: str) -> List[Tuple[str, str, bool, bool]]:
        """返回表结构：[(列名, 类型, 是否可空, 是否主键)]"""
        rows = self.fetchall(f"PRAGMA table_info('{table_name}');")
        schema = []
        for row in rows:
            col_name = row[1]
            col_type = row[2]
            is_nullable = row[3] == 0
            is_primary = row[5] > 0
            schema.append((col_name, col_type, is_nullable, is_primary))
        return schema

    def get_row_count(self, table_name: str) -> int:
        try:
            result = self.fetchone(f"SELECT COUNT(*) FROM '{table_name}';")
            return result[0] if result else 0
        except Exception:
            return 0

    def get_columns(self, table_name: str) -> List[str]:
        rows = self.fetchall(f"PRAGMA table_info('{table_name}');")
        return [row[1] for row in rows]


class MySQLConnector(DatabaseConnector):
    """MySQL 数据库连接器"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__("mysql", config)
        self.host = config.get("host", "127.0.0.1")
        self.port = int(config.get("port", 3306))
        self.user = config.get("user", "root")
        self.password = str(config.get("password", ""))
        self.database = config.get("database")
        self.charset = config.get("charset", "utf8mb4")

    def connect(self):
        import pymysql
        self.connection = pymysql.connect(
            host=self.host,
            port=self.port,
            user=self.user,
            password=self.password,
            database=self.database if self.database else None,
            charset=self.charset,
            connect_timeout=30,
            cursorclass=pymysql.cursors.DictCursor,
        )
        return self.connection

    def get_databases(self) -> List[str]:
        rows = self.fetchall("SHOW DATABASES;")
        return [row["Database"] if isinstance(row, dict) else row[0]
                for row in rows
                if row and (row.get("Database") if isinstance(row, dict) else row[0])
                not in ("information_schema", "mysql", "performance_schema", "sys")]

    def get_tables(self, database: Optional[str] = None) -> List[str]:
        if database:
            self.execute(f"USE `{database}`;")
        rows = self.fetchall("SHOW TABLES;")
        tables = []
        for row in rows:
            if isinstance(row, dict):
                tables.append(list(row.values())[0])
            else:
                tables.append(row[0])
        return tables

    def get_table_schema(self, table_name: str) -> List[Tuple[str, str, bool, bool]]:
        rows = self.fetchall(f"DESCRIBE `{table_name}`;")
        schema = []
        for row in rows:
            if isinstance(row, dict):
                col_name = row["Field"]
                col_type = row["Type"]
                is_nullable = row["Null"].upper() == "YES"
                is_primary = "PRI" in (row.get("Key") or "")
            else:
                col_name = row[0]
                col_type = row[1]
                is_nullable = row[2].upper() == "YES"
                # DESCRIBE 顺序: Field(0), Type(1), Null(2), Key(3), Default(4), Extra(5)
                is_primary = "PRI" in (row[3] or "")
            schema.append((col_name, col_type, is_nullable, is_primary))
        return schema

    def get_row_count(self, table_name: str) -> int:
        try:
            result = self.fetchone(f"SELECT COUNT(*) AS cnt FROM `{table_name}`;")
            if isinstance(result, dict):
                return int(result.get("cnt") or 0)
            return result[0] if result else 0
        except Exception:
            return 0

    def streaming_fetchall(self, columns: List[str], table: str,
                            batch_size: int = 50000) -> Any:
        """MySQL 流式读取（SSCursor，服务端游标）- 速度比 LIMIT/OFFSET 快 5-10 倍。

        返回一个生成器，每次 yield 一批数据（List[tuple]）。
        整个过程只需要 1 次 SQL 查询，避免多次网络往返。
        使用 fetchmany 代替 fetchone，减少 Python 循环调用次数。

        注意：仅 MySQL 有效，其他数据库返回 None。
        """
        try:
            import pymysql
        except Exception:
            return None

        # 建立独立的 SSCursor 连接（不能和 DictCursor 共享）
        try:
            streaming_conn = pymysql.connect(
                host=self.host,
                port=self.port,
                user=self.user,
                password=self.password,
                database=self.database if self.database else None,
                charset=self.charset,
                connect_timeout=600,
                cursorclass=pymysql.cursors.SSCursor,  # 关键：服务端流式游标
            )
        except Exception:
            return None

        stream_cursor = streaming_conn.cursor()
        col_placeholders = ", ".join(["`" + c + "`" for c in columns])

        def _generator():
            try:
                stream_cursor.execute(f"SELECT {col_placeholders} FROM `{table}`")
                while True:
                    # fetchmany 一次取 batch_size 行，比逐行 fetchone 快得多
                    rows = stream_cursor.fetchmany(batch_size)
                    if not rows:
                        break
                    yield [tuple(r) for r in rows]
            finally:
                # 防御式清理：先断开 cursor 内部引用，再关闭 cursor，最后关 connection
                # 避免用户中断时 __del__ 访问已关闭 socket
                try:
                    if hasattr(stream_cursor, '_result'):
                        try:
                            if stream_cursor._result is not None and hasattr(stream_cursor._result, 'connection'):
                                stream_cursor._result.connection = None
                        except Exception: pass
                    stream_cursor.close()
                except Exception: pass
                try:
                    if hasattr(streaming_conn, '_sock'):
                        try: streaming_conn._sock = None
                        except Exception: pass
                    streaming_conn.close()
                except Exception: pass

        return _generator()

    def get_columns(self, table_name: str) -> List[str]:
        schema = self.get_table_schema(table_name)
        return [col[0] for col in schema]


class PostgreSQLConnector(DatabaseConnector):
    """PostgreSQL 数据库连接器"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__("postgresql", config)
        self.host = config.get("host", "127.0.0.1")
        self.port = int(config.get("port", 5432))
        self.user = config.get("user", "postgres")
        self.password = str(config.get("password", ""))
        self.database = config.get("database", "postgres")

    def connect(self):
        import psycopg2
        self.connection = psycopg2.connect(
            host=self.host,
            port=self.port,
            user=self.user,
            password=self.password,
            dbname=self.database,
            connect_timeout=30,
        )
        self.connection.autocommit = False
        return self.connection

    def get_databases(self) -> List[str]:
        rows = self.fetchall(
            "SELECT datname FROM pg_database WHERE datistemplate = FALSE ORDER BY datname;"
        )
        exclude = {"postgres", "template0", "template1"}
        return [row[0] for row in rows if row[0] not in exclude]

    def get_tables(self, database: Optional[str] = None) -> List[str]:
        rows = self.fetchall(
            "SELECT table_name FROM information_schema.tables "
            "WHERE table_schema = 'public' ORDER BY table_name;"
        )
        return [row[0] for row in rows]

    def get_table_schema(self, table_name: str) -> List[Tuple[str, str, bool, bool]]:
        sql = (
            "SELECT c.column_name, c.data_type, c.is_nullable, "
            "EXISTS(SELECT 1 FROM information_schema.key_column_usage kcu "
            "JOIN information_schema.table_constraints tc "
            "ON kcu.constraint_name = tc.constraint_name "
            "WHERE tc.constraint_type = 'PRIMARY KEY' "
            "AND kcu.table_name = c.table_name AND kcu.column_name = c.column_name) AS is_pk "
            "FROM information_schema.columns c "
            "WHERE c.table_name = %s AND c.table_schema = 'public' "
            "ORDER BY c.ordinal_position;"
        )
        rows = self.fetchall(sql, (table_name,))
        schema = []
        for row in rows:
            col_name = row[0]
            col_type = row[1]
            is_nullable = (row[2] or "").upper() == "YES"
            is_primary = bool(row[3])
            schema.append((col_name, col_type, is_nullable, is_primary))
        return schema

    def get_row_count(self, table_name: str) -> int:
        try:
            result = self.fetchone(
                f'SELECT COUNT(*) FROM "{table_name}";'
            )
            return result[0] if result else 0
        except Exception:
            return 0

    def get_columns(self, table_name: str) -> List[str]:
        schema = self.get_table_schema(table_name)
        return [col[0] for col in schema]


def create_connector(db_type: str, config: Dict[str, Any]) -> DatabaseConnector:
    """工厂方法：根据数据库类型创建对应的连接器"""
    db_type_lower = db_type.lower().strip()
    if db_type_lower == "sqlite":
        return SQLiteConnector(config)
    if db_type_lower == "mysql":
        return MySQLConnector(config)
    if db_type_lower in ("postgresql", "psql", "pg"):
        return PostgreSQLConnector(config)
    raise ValueError(f"未知的数据库类型: {db_type}")
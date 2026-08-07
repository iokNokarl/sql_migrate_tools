"""
数据库导出命令行接口
使用子命令 `export` 触发导出功能。
"""

import argparse
import os
import sys
import time
from datetime import datetime
from typing import Any, Dict, List, Optional

from .database_connector import create_connector
from .db_exporter import (
    BATCH_THRESHOLD,
    DEFAULT_BATCH_SIZE,
    MAX_ROWS_PER_SHEET,
    MAX_ROWS_PER_WORKBOOK,
    MULTITHREAD_THRESHOLD,
    DEFAULT_THREADS,
    MAX_THREADS,
    DatabaseExporter,
    HAS_OPENPYXL,
    HAS_XLSXWRITER,
)


def _parse_comma_list(value: Optional[str]) -> Optional[List[str]]:
    """解析逗号分隔列表"""
    if not value or not value.strip():
        return None
    return [item.strip() for item in value.split(",") if item.strip()]


def create_export_parser(subparsers: Optional[argparse._SubParsersAction] = None) -> argparse.ArgumentParser:
    """
    创建导出参数解析器。
    支持作为子命令（subparsers 传入时）或独立解析器使用。
    """
    parser_kwargs: Dict[str, Any] = {
        "prog": "run.py export",
        "description": (
            "数据库导出工具 - 支持将 SQLite / MySQL / PostgreSQL 中的表 "
            "导出为 Excel(.xlsx) 或 CSV 文件\n"
            f"特性：超过 {BATCH_THRESHOLD} 条自动分批（每批 {DEFAULT_BATCH_SIZE} 条），"
            f"单 sheet 最多 {MAX_ROWS_PER_SHEET} 行，单工作簿合计最多 {MAX_ROWS_PER_WORKBOOK} 行"
        ),
        "formatter_class": argparse.RawDescriptionHelpFormatter,
        "epilog": (
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "使用示例：\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "\n"
            "【1】导出 SQLite 全部表为 Excel（默认）\n"
            "  python run.py export --source-type sqlite --source-path ./data/source.db \\\n"
            "      --format xlsx --output-dir ./exports -y\n"
            "\n"
            "【2】导出 MySQL 指定表为 CSV（自定义文件名）\n"
            "  python run.py export --source-type mysql --source-host 127.0.0.1 \\\n"
            "      --source-port 3306 --source-user root --source-password 123456 \\\n"
            "      --source-database test_db --tables users,orders --format csv \\\n"
            "      --filename my_export -y\n"
            "\n"
            "【3】导出 PostgreSQL 指定行范围（第 100~600 行）\n"
            "  python run.py export --source-type postgresql --source-host 127.0.0.1 \\\n"
            "      --source-port 5432 --source-user postgres --source-password 123456 \\\n"
            "      --source-database test_db --tables products --from-row 100 --to-row 600 -y\n"
            "\n"
            "【4】导出只取前 500 行（limit 用法）\n"
            "  python run.py export --source-type sqlite --source-path ./data/source.db \\\n"
            "      --tables employees --limit 500 -y\n"
            "\n"
            "【5】自定义每批大小（适合超大数据量场景）\n"
            "  python run.py export --source-type sqlite --source-path ./data/source.db \\\n"
            "      --tables huge_table --batch-size 1000 -y\n"
        ),
    }

    if subparsers is not None:
        parser = subparsers.add_parser("export", **parser_kwargs)
    else:
        parser = argparse.ArgumentParser(**parser_kwargs)

    # 源数据库参数
    parser.add_argument(
        "--source-type",
        required=True,
        choices=["sqlite", "mysql", "postgresql", "psql", "pg"],
        help="源数据库类型",
    )
    parser.add_argument("--source-host", default="127.0.0.1", help="源数据库主机地址")
    parser.add_argument("--source-port", type=int, default=0, help="源数据库端口号（0 表示使用默认端口）")
    parser.add_argument("--source-user", default="", help="源数据库用户名")
    parser.add_argument("--source-password", default="", help="源数据库密码")
    parser.add_argument("--source-database", default="", help="源数据库名称")
    parser.add_argument("--source-path", default="", help="SQLite 数据库文件路径")
    parser.add_argument("--source-charset", default="utf8mb4", help="MySQL 字符集")

    # 导出参数
    parser.add_argument(
        "--tables",
        type=_parse_comma_list,
        default=None,
        help="指定要导出的表名（逗号分隔，如: users,orders），不指定则导出全部表",
    )
    parser.add_argument(
        "--format",
        "--output-format",
        dest="output_format",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="导出格式（默认 xlsx）",
    )
    parser.add_argument(
        "--output-dir",
        default="./exports",
        help="输出目录（默认 ./exports）",
    )
    parser.add_argument(
        "--filename",
        default=None,
        help="自定义文件名（不含扩展名）。留空则自动生成: 数据库名_表名_行数_年月日_时分秒",
    )
    parser.add_argument(
        "--from-row",
        type=int,
        default=0,
        help="从第几行开始导出（从 0 开始计数，默认 0）",
    )
    parser.add_argument(
        "--to-row",
        type=int,
        default=None,
        help="导出到第几行（不含该行，默认到最后）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="最多导出多少行（优先级低于 --to-row）",
    )
    parser.add_argument(
        "--threads",
        type=int,
        default=DEFAULT_THREADS,
        help=f"最大线程数（默认 {DEFAULT_THREADS}，最多 {MAX_THREADS}，超过 {MULTITHREAD_THRESHOLD} 行时启用多线程",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=DEFAULT_BATCH_SIZE,
        help=f"分批读取时每批行数（默认 {DEFAULT_BATCH_SIZE}），超过 {BATCH_THRESHOLD} 条时生效",
    )
    parser.add_argument(
        "-y", "--yes",
        action="store_true",
        default=False,
        help="跳过导出前的二次确认",
    )

    return parser


def _build_source_config(args: argparse.Namespace) -> Dict[str, Any]:
    """根据参数构造源数据库配置字典"""
    config: Dict[str, Any] = {}
    db_type = args.source_type.lower()

    if db_type == "sqlite":
        config["path"] = args.source_path or args.source_database
        if not config["path"]:
            raise ValueError("SQLite 必须通过 --source-path 指定数据库文件路径")
    elif db_type == "mysql":
        config["host"] = args.source_host or "127.0.0.1"
        config["port"] = int(args.source_port or 3306)
        config["user"] = args.source_user or "root"
        config["password"] = str(args.source_password or "")
        config["database"] = args.source_database
        config["charset"] = args.source_charset or "utf8mb4"
        if not config["database"]:
            raise ValueError("MySQL 必须通过 --source-database 指定数据库名")
    elif db_type in ("postgresql", "psql", "pg"):
        config["host"] = args.source_host or "127.0.0.1"
        config["port"] = int(args.source_port or 5432)
        config["user"] = args.source_user or "postgres"
        config["password"] = str(args.source_password or "")
        config["database"] = args.source_database
        if not config["database"]:
            raise ValueError("PostgreSQL 必须通过 --source-database 指定数据库名")

    return config


def _print_summary(
    args: argparse.Namespace,
    tables: List[str],
    db_display: str,
) -> None:
    """打印导出摘要"""
    print("=" * 60)
    print("【导出计划摘要】")
    print("-" * 60)
    print(f"源数据库: {args.source_type.upper()}")
    if args.source_type.lower() == "sqlite":
        print(f"  文件: {args.source_path or args.source_database}")
    else:
        print(f"  主机: {args.source_host}:{args.source_port or '(默认)'}")
        print(f"  库名: {args.source_database}")
    print(f"导出格式: {args.output_format.upper()}")
    print(f"输出目录: {os.path.abspath(args.output_dir)}")
    if args.filename:
        print(f"自定义文件名: {args.filename}.{args.output_format}")
    else:
        print("文件名规则: 数据库名_表名_行数_年月日_时分秒")

    range_info = []
    if args.from_row > 0:
        range_info.append(f"从第 {args.from_row} 行起")
    if args.to_row is not None:
        range_info.append(f"到第 {args.to_row} 行止")
    if args.limit is not None:
        range_info.append(f"最多 {args.limit} 行")
    if range_info:
        print(f"行范围: {'，'.join(range_info)}")

    print(f"待导出表: {'，'.join(tables)}（共 {len(tables)} 张）")
    print(f"分批大小: {args.batch_size} 行/批（超过 {BATCH_THRESHOLD} 条启用）")
    print("=" * 60)
    print()


def _confirm(prompt: str) -> bool:
    """简单的 y/N 确认"""
    try:
        answer = input(f"{prompt} (y/N): ").strip().lower()
    except EOFError:
        return False
    return answer in ("y", "yes")


def run_export(args: Optional[argparse.Namespace] = None) -> int:
    """
    执行导出流程。
    传入 args 时使用该参数对象，否则从命令行解析。
    返回退出码。
    """
    if args is None:
        parser = create_export_parser()
        args = parser.parse_args()

    # 1. 驱动检查
    if args.output_format == "xlsx" and not (HAS_XLSXWRITER or HAS_OPENPYXL):
        print("[ERROR] 导出 Excel 需要安装 xlsxwriter 或 openpyxl")
        print("        请执行: pip install xlsxwriter")
        return 1

    source_db_type = args.source_type.lower()
    if source_db_type == "mysql":
        try:
            import pymysql  # noqa: F401
        except ImportError:
            print("[ERROR] 连接 MySQL 需要安装 pymysql 包")
            print("        请执行: pip install pymysql")
            return 1
    elif source_db_type in ("postgresql", "psql", "pg"):
        try:
            import psycopg2  # noqa: F401
        except ImportError:
            print("[ERROR] 连接 PostgreSQL 需要安装 psycopg2-binary 包")
            print("        请执行: pip install psycopg2-binary")
            return 1

    # 2. 构造配置
    try:
        source_config = _build_source_config(args)
    except ValueError as exc:
        print(f"[ERROR] 参数配置错误: {exc}")
        return 2

    # 3. 连接数据库并获取表
    try:
        connector = create_connector(args.source_type, source_config)
        connector.connect()
    except Exception as exc:
        print(f"[ERROR] 数据库连接失败: {exc}")
        return 4

    try:
        all_tables = connector.get_tables(source_config.get("database"))
        tables = args.tables if args.tables else all_tables

        if not tables:
            print("[WARN] 未找到任何可导出的表")
            return 0

        # 校验指定的表是否存在
        if args.tables:
            missing = [t for t in args.tables if t not in all_tables]
            if missing:
                print(f"[ERROR] 指定的表不存在: {', '.join(missing)}")
                print(f"        可用的表: {', '.join(all_tables)}")
                return 2

        # 4. 打印摘要与确认
        db_display = source_config.get("database") or (
            os.path.splitext(os.path.basename(source_config.get("path", "")))[0]
            if source_config.get("path")
            else args.source_type
        )
        _print_summary(args, tables, db_display)

        if not args.yes:
            if not _confirm("是否确认开始导出？"):
                print("[INFO] 已取消导出")
                return 0

        # 5. 执行导出
        exporter = DatabaseExporter(
            connector=connector,
            tables=tables,
            db_type=args.source_type,
            db_config=source_config,
            output_format=args.output_format,
            output_dir=args.output_dir,
            custom_filename=args.filename,
            from_row=args.from_row,
            to_row=args.to_row,
            limit=args.limit,
            batch_size=args.batch_size,
            max_threads=args.threads,
            logger=None,
        )

        summary = exporter.export()

        # 6. 打印结果汇总
        elapsed = summary.get("elapsed", 0.0)
        print()
        print("=" * 60)
        print("【导出结果汇总】")
        print("-" * 60)
        for tname, tinfo in summary["tables"].items():
            files_str = ", ".join(os.path.basename(f) for f in tinfo["files"])
            print(f"  [OK] {tname}: {tinfo['rows']} 行 / {tinfo['columns']} 列 / {len(tinfo['files'])} 个文件")
            print(f"       文件: {files_str}")
        print("-" * 60)
        print(f"总计: {len(summary['tables'])} 张表 / {summary['total_rows']} 行 / "
              f"{len(summary['output_files'])} 个文件")
        print(f"总耗时: {elapsed:.2f} 秒")
        print(f"输出目录: {os.path.abspath(args.output_dir)}")
        print("=" * 60)

        return 0

    except KeyboardInterrupt:
        print("\n[INFO] 用户中断操作")
        return 130
    except Exception as exc:
        print(f"[ERROR] 导出失败: {exc}")
        import traceback
        traceback.print_exc()
        return 5
    finally:
        try:
            connector.close()
        except Exception:
            pass


def main() -> int:
    """独立入口"""
    return run_export()


if __name__ == "__main__":
    sys.exit(main())
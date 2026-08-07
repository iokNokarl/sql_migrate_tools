"""
数据库迁移工具 - 命令行入口
支持 SQLite / MySQL / PostgreSQL 三种数据库之间的互迁移

使用示例：
    # 从 SQLite 迁移到 MySQL（迁移所有表）
    python db_migration_tool.py \
        --source-type sqlite --source-path ./data/source.db \
        --target-type mysql --target-host 127.0.0.1 --target-port 3306 \
        --target-user root --target-password 123456 --target-database target_db

    # 从 PostgreSQL 迁移到 SQLite（迁移指定表）
    python db_migration_tool.py \
        --source-type postgresql --source-host 127.0.0.1 --source-port 5432 \
        --source-user postgres --source-password 123456 --source-database source_db \
        --target-type sqlite --target-path ./data/target.db \
        --tables users,orders,products --threads 8

    # 迁移指定库（仅适用于支持多库的源数据库）
    python db_migration_tool.py \
        --source-type mysql --source-host 127.0.0.1 --source-user root \
        --source-password 123456 --source-database db1 \
        --target-type mysql --target-host 127.0.0.1 --target-user root \
        --target-password 123456 --target-database db2 \
        --databases db1
"""

import argparse
import os
import sys
import traceback
from typing import Dict, List, Optional


REQUIRED_LIBRARIES = {
    "sqlite": {"module": "sqlite3", "install": "（Python 标准库，无需安装）"},
    "mysql": {"module": "pymysql", "install": "pip install pymysql"},
    "postgresql": {"module": "psycopg2", "install": "pip install psycopg2-binary"},
    "psql": {"module": "psycopg2", "install": "pip install psycopg2-binary"},
    "pg": {"module": "psycopg2", "install": "pip install psycopg2-binary"},
}


def check_libraries(db_type: str) -> bool:
    """检查数据库驱动是否已安装"""
    db_type_lower = db_type.lower()
    if db_type_lower not in REQUIRED_LIBRARIES:
        print(f"[ERROR] 未知数据库类型: {db_type}")
        return False

    lib_info = REQUIRED_LIBRARIES[db_type_lower]
    module_name = lib_info["module"]
    try:
        __import__(module_name)
        print(f"[INFO] 数据库驱动 {module_name} 已安装 ✓")
        return True
    except ImportError:
        print(f"[ERROR] 缺少数据库驱动: {module_name}")
        print(f"        请执行: {lib_info['install']}")
        return False


def check_all_libraries(source_type: str, target_type: str) -> bool:
    """检查源和目标数据库都需要的驱动"""
    print("=" * 60)
    print("正在检查 Python 数据库驱动依赖...")
    source_ok = check_libraries(source_type)
    target_ok = check_libraries(target_type)
    print("=" * 60)
    return source_ok and target_ok


def build_source_config(args: argparse.Namespace) -> Dict:
    """构建源数据库配置"""
    db_type = args.source_type.lower()
    if db_type == "sqlite":
        if not args.source_path:
            raise ValueError("SQLite 源数据库必须指定 --source-path")
        return {"path": args.source_path}
    if db_type == "mysql":
        return {
            "host": args.source_host,
            "port": args.source_port,
            "user": args.source_user,
            "password": args.source_password or "",
            "database": args.source_database,
            "charset": args.source_charset,
        }
    if db_type in ("postgresql", "psql", "pg"):
        return {
            "host": args.source_host,
            "port": args.source_port,
            "user": args.source_user,
            "password": args.source_password or "",
            "database": args.source_database or "postgres",
        }
    raise ValueError(f"不支持的源数据库类型: {db_type}")


def build_target_config(args: argparse.Namespace) -> Dict:
    """构建目标数据库配置"""
    db_type = args.target_type.lower()
    if db_type == "sqlite":
        if not args.target_path:
            raise ValueError("SQLite 目标数据库必须指定 --target-path")
        parent_dir = os.path.dirname(os.path.abspath(args.target_path))
        if parent_dir and not os.path.exists(parent_dir):
            os.makedirs(parent_dir, exist_ok=True)
        return {"path": args.target_path}
    if db_type == "mysql":
        return {
            "host": args.target_host,
            "port": args.target_port,
            "user": args.target_user,
            "password": args.target_password or "",
            "database": args.target_database,
            "charset": args.target_charset,
        }
    if db_type in ("postgresql", "psql", "pg"):
        return {
            "host": args.target_host,
            "port": args.target_port,
            "user": args.target_user,
            "password": args.target_password or "",
            "database": args.target_database or "postgres",
        }
    raise ValueError(f"不支持的目标数据库类型: {db_type}")


def parse_comma_separated(value: Optional[str]) -> Optional[List[str]]:
    """解析逗号分隔参数"""
    if not value:
        return None
    return [item.strip() for item in value.split(",") if item.strip()]


def parse_table_rename(value: Optional[str]) -> Optional[Dict[str, str]]:
    """解析表名映射参数，格式: 源表名:目标表名,源表2:目标表2"""
    if not value:
        return None
    mapping: Dict[str, str] = {}
    for pair in value.split(","):
        pair = pair.strip()
        if not pair or ":" not in pair:
            raise ValueError(
                f"表名映射格式错误：{pair}，正确格式: 源表名:目标表名，如 users:t_users"
            )
        src, tgt = pair.split(":", 1)
        src = src.strip()
        tgt = tgt.strip()
        if not src or not tgt:
            raise ValueError(f"表名映射不能为空：{pair}")
        mapping[src] = tgt
    return mapping


def create_arg_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        prog="run.py",
        description="数据库迁移工具 - 支持 SQLite / MySQL / PostgreSQL 三者互迁移\n"
                    "功能：分批处理 / 多线程 / 事务回滚 / 指定表或库迁移",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "使用示例：\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "\n"
            "【1】SQLite 迁移到 MySQL（迁移全部表）\n"
            "  python run.py \\\n"
            "      --source-type sqlite --source-path ./data/source.db \\\n"
            "      --target-type mysql --target-host 127.0.0.1 --target-port 3306 \\\n"
            "      --target-user root --target-password 123456 --target-database target_db -y\n"
            "\n"
            "【2】MySQL 迁移到 PostgreSQL（指定表，启用多线程）\n"
            "  python run.py \\\n"
            "      --source-type mysql --source-host 127.0.0.1 --source-port 3306 \\\n"
            "      --source-user root --source-password 123456 --source-database src_db \\\n"
            "      --target-type postgresql --target-host 127.0.0.1 --target-port 5432 \\\n"
            "      --target-user postgres --target-password 123456 --target-database target_db \\\n"
            "      --tables users,orders,products --threads 10 --batch-size 10000 -y\n"
            "\n"
            "【3】PostgreSQL 迁移到 SQLite（只迁移一个表，小批量）\n"
            "  python run.py \\\n"
            "      --source-type postgresql --source-host 127.0.0.1 --source-port 5432 \\\n"
            "      --source-user postgres --source-password 123456 --source-database src_db \\\n"
            "      --target-type sqlite --target-path ./backup/export.db \\\n"
            "      --tables employees --batch-size 1000 -y\n"
            "\n"
            "【4】SQLite 之间迁移（用于测试、备份）\n"
            "  python run.py \\\n"
            "      --source-type sqlite --source-path ./data/src.db \\\n"
            "      --target-type sqlite --target-path ./data/backup.db -y\n"
            "\n"
            "参数说明：\n"
            "  • -y                 跳过迁移前的确认提示\n"
            "  • --tables           指定要迁移的表名（逗号分隔），不填则迁移全部表\n"
            "  • --rename           表名映射，源表名:目标表名（逗号分隔多对，如: users:t_users,orders:t_orders）\n"
            "  • --databases        指定要迁移的数据库（MySQL 支持多库）\n"
            "  • --threads          单表数据超过 5 万条时启用的最大线程数（默认 8，最多 20）\n"
            "  • --batch-size       每批处理条数（默认 10000），超过 3000 条自动启用分批\n"
        ),
    )

    parser.add_argument(
        "--source-type",
        required=True,
        choices=["sqlite", "mysql", "postgresql", "psql", "pg"],
        help="源数据库类型",
    )
    parser.add_argument("--source-host", default="127.0.0.1", help="源数据库主机")
    parser.add_argument("--source-port", type=int, default=0, help="源数据库端口")
    parser.add_argument("--source-user", default="", help="源数据库用户名")
    parser.add_argument("--source-password", default="", help="源数据库密码")
    parser.add_argument("--source-database", default="", help="源数据库名")
    parser.add_argument("--source-path", default="", help="SQLite 源数据库文件路径")
    parser.add_argument("--source-charset", default="utf8mb4", help="MySQL 源数据库字符集")

    parser.add_argument(
        "--target-type",
        required=True,
        choices=["sqlite", "mysql", "postgresql", "psql", "pg"],
        help="目标数据库类型",
    )
    parser.add_argument("--target-host", default="127.0.0.1", help="目标数据库主机")
    parser.add_argument("--target-port", type=int, default=0, help="目标数据库端口")
    parser.add_argument("--target-user", default="", help="目标数据库用户名")
    parser.add_argument("--target-password", default="", help="目标数据库密码")
    parser.add_argument("--target-database", default="", help="目标数据库名")
    parser.add_argument("--target-path", default="", help="SQLite 目标数据库文件路径")
    parser.add_argument("--target-charset", default="utf8mb4", help="MySQL 目标数据库字符集")

    parser.add_argument(
        "--tables",
        default=None,
        help="指定要迁移的表名，多个表用逗号分隔（如: users,orders），不指定则迁移全部表",
    )
    parser.add_argument(
        "--rename",
        default=None,
        help="表名映射（源表:目标表，逗号分隔多对），如: users:t_users,orders:t_orders",
    )
    parser.add_argument(
        "--databases",
        default=None,
        help="指定要迁移的数据库名（仅适用于支持多库的数据库），多个库用逗号分隔",
    )
    parser.add_argument(
        "--threads",
        type=int,
        default=8,
        help="异步处理时的最大线程数（默认 8，最多 20），超过 5 万条数据时启用",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=10000,
        help="每批处理条数（默认 10000），超过 3000 条数据时启用分批",
    )
    parser.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="跳过迁移前的二次确认",
    )
    return parser


def display_migration_plan(
    args: argparse.Namespace,
    databases: Optional[List[str]],
    tables: Optional[List[str]],
    table_rename: Optional[Dict[str, str]] = None,
) -> None:
    """展示迁移计划摘要"""
    print("\n" + "=" * 60)
    print("【迁移计划摘要】")
    print("-" * 60)
    print(f"源数据库: {args.source_type.upper()}")
    if args.source_type.lower() == "sqlite":
        print(f"  路径: {args.source_path}")
    else:
        print(f"  主机: {args.source_host}:{args.source_port or '(默认)'}")
        print(f"  用户: {args.source_user or '(空)'}")
        print(f"  数据库: {args.source_database or '(未指定)'}")
    print(f"目标数据库: {args.target_type.upper()}")
    if args.target_type.lower() == "sqlite":
        print(f"  路径: {args.target_path}")
    else:
        print(f"  主机: {args.target_host}:{args.target_port or '(默认)'}")
        print(f"  用户: {args.target_user or '(空)'}")
        print(f"  数据库: {args.target_database or '(未指定)'}")
    if tables:
        print(f"指定表: {', '.join(tables)}（共 {len(tables)} 张）")
    if table_rename:
        rename_lines = [f"{k} → {v}" for k, v in table_rename.items()]
        print(f"表名映射: {', '.join(rename_lines)}")
    if databases:
        print(f"指定数据库: {', '.join(databases)}（共 {len(databases)} 个）")
    print(f"最大线程数: {args.threads}")
    print(f"分批大小: {args.batch_size}")
    print("=" * 60 + "\n")


def confirm_migration(yes_flag: bool) -> bool:
    """交互式确认迁移"""
    if yes_flag:
        return True
    try:
        answer = input("请确认是否开始迁移？(y/N) ").strip().lower()
        return answer in ("y", "yes")
    except (EOFError, KeyboardInterrupt):
        print("\n[INFO] 用户取消迁移")
        return False


def run_migration(args: argparse.Namespace) -> int:
    """执行主迁移流程"""
    try:
        source_config = build_source_config(args)
        target_config = build_target_config(args)
    except ValueError as exc:
        print(f"[ERROR] 配置错误: {exc}")
        return 2

    databases = parse_comma_separated(args.databases)
    tables = parse_comma_separated(args.tables)
    max_threads = min(max(1, args.threads), 20)
    batch_size = max(100, args.batch_size)
    # 解析表名映射
    table_rename = parse_table_rename(args.rename)

    display_migration_plan(args, databases, tables, table_rename)

    if not confirm_migration(args.yes):
        print("[INFO] 迁移已取消")
        return 0

    try:
        from .database_connector import create_connector
        from .database_migrator import (
            DatabaseMigrator,
            MigrationError,
            MigrationLogger,
        )
    except Exception as exc:
        print(f"[ERROR] 加载迁移模块失败: {exc}")
        traceback.print_exc()
        return 3

    source_connector = None
    target_connector = None
    try:
        source_connector = create_connector(args.source_type, source_config)
        target_connector = create_connector(args.target_type, target_config)

        MigrationLogger.info("正在测试源数据库连接...")
        source_connector.connect()
        source_connector.close()
        MigrationLogger.success("源数据库连接正常")

        MigrationLogger.info("正在测试目标数据库连接...")
        target_connector.connect()
        target_connector.close()
        MigrationLogger.success("目标数据库连接正常")
    except Exception as exc:
        print(f"[ERROR] 数据库连接失败: {exc}")
        MigrationLogger.error("请检查数据库地址、端口、账号、权限是否正确")
        return 4
    finally:
        if source_connector:
            try:
                source_connector.close()
            except Exception:
                pass
        if target_connector:
            try:
                target_connector.close()
            except Exception:
                pass

    migrator = DatabaseMigrator(
        source_connector=create_connector(args.source_type, source_config),
        target_connector=create_connector(args.target_type, target_config),
        max_threads=max_threads,
        batch_size=batch_size,
    )

    try:
        result = migrator.migrate(databases=databases, tables=tables, table_rename=table_rename)
        print("\n" + "=" * 60)
        print("【迁移结果汇总】")
        print("-" * 60)
        for table_result in result["tables"]:
            db_info = f"[{table_result['database']}] " if table_result["database"] else ""
            src = table_result["table"]
            tgt = table_result.get("target_table")
            display = src if (tgt is None or tgt == src) else f"{src} → {tgt}"
            print(
                f"  ✓ {db_info}{display}: "
                f"{table_result['row_count']} 行，"
                f"{table_result['batch_count'] or 1} 批，"
                f"{table_result['thread_count']} 线程，"
                f"耗时 {table_result['elapsed']:.2f}s"
            )
        print("-" * 60)
        print(f"总计: {len(result['tables'])} 张表，{result['total_rows']} 条记录")
        print(f"总耗时: {result['elapsed']:.2f} 秒")
        print("=" * 60)
        return 0
    except MigrationError as exc:
        print(f"\n[ERROR] 迁移失败: {exc}")
        if exc.table:
            print(f"        失败位置: 表 {exc.table}")
        if exc.database:
            print(f"        失败数据库: {exc.database}")
        return 5
    except Exception as exc:
        print(f"\n[ERROR] 发生未预期错误: {exc}")
        traceback.print_exc()
        return 99


def main() -> int:
    """主入口"""
    parser = create_arg_parser()
    args = parser.parse_args()

    if not check_all_libraries(args.source_type, args.target_type):
        print("\n[ERROR] 请先安装缺少的数据库驱动后重试")
        return 1

    return run_migration(args)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n[INFO] 用户中断操作")
        sys.exit(130)
    except Exception as exc:
        print(f"\n[FATAL] 程序崩溃: {exc}")
        traceback.print_exc()
        sys.exit(2)
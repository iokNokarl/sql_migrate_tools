"""
数据迁移/导出工具 - 启动脚本
支持两个子命令：
  migrate - 数据库迁移（SQLite ↔ MySQL ↔ PostgreSQL）
  export  - 将数据库表导出为 Excel(.xlsx) 或 CSV

若省略子命令，默认进入 migrate 模式（向后兼容）。
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def _build_parser():
    """用 argparse 构建完整的子命令解析器，保证 -h 信息规范、详细"""
    import argparse

    parser = argparse.ArgumentParser(
        prog="run.py",
        description="数据工具 - 支持数据库迁移(migrate) 与 Excel/CSV 导出(export)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "查看各子命令详细参数:\n"
            "  python run.py migrate -h\n"
            "  python run.py export  -h\n"
            "\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "【快速示例】\n"
            "\n"
            "① 迁移 SQLite → MySQL（全部表）\n"
            "  python run.py migrate --source-type sqlite --source-path ./src.db \\\n"
            "      --target-type mysql --target-host 127.0.0.1 --target-port 3306 \\\n"
            "      --target-user root --target-password 123456 --target-database target_db -y\n"
            "\n"
            "② 导出 SQLite 指定表为 Excel\n"
            "  python run.py export --source-type sqlite --source-path ./data/source.db \\\n"
            "      --tables users,orders --format xlsx --output-dir ./exports -y\n"
            "\n"
            "③ 导出 PostgreSQL 指定行范围为 CSV\n"
            "  python run.py export --source-type postgresql --source-host 127.0.0.1 \\\n"
            "      --source-port 5432 --source-user postgres --source-password 123456 \\\n"
            "      --source-database test_db --tables products --from-row 100 --limit 500 -y\n"
            "\n"
            "④ 大数量场景（超过 50 万行自动分批，超过 50 万行自动分文件）\n"
            "  python run.py export --source-type mysql --source-host 127.0.0.1 \\\n"
            "      --source-user root --source-password 123456 --source-database big_db \\\n"
            "      --tables huge_table --batch-size 2500 --threads 8 -y\n"
        ),
    )

    subparsers = parser.add_subparsers(
        dest="subcommand",
        metavar="<子命令>",
        title="子命令",
        description=(
            "migrate : 数据库迁移（SQLite ↔ MySQL ↔ PostgreSQL 三者互迁移，\n"
            "          支持分批处理、多线程、事务回滚）\n"
            "export  : 将数据库表导出为 Excel(.xlsx) 或 CSV，\n"
            "          支持分批读取、多线程、分 sheet / 分文件 / 子文件夹"
        ),
    )

    # ---------------- migrate 子命令 ----------------
    p_migrate = subparsers.add_parser(
        "migrate",
        help="数据库迁移（SQLite ↔ MySQL ↔ PostgreSQL 三者互迁移）",
        description="数据库迁移 - 支持 SQLite / MySQL / PostgreSQL 三者互迁移，"
                    "超过 3000 行自动分批（每批 2500 条），超过 5 万行启用多线程。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    # 源数据库
    s_src = p_migrate.add_argument_group("源数据库配置")
    s_src.add_argument("--source-type", required=True,
                        choices=["sqlite", "mysql", "postgresql", "pg", "psql"],
                        help="源数据库类型")
    s_src.add_argument("--source-host", default="127.0.0.1",
                        help="源数据库主机地址 (默认 127.0.0.1)")
    s_src.add_argument("--source-port", type=int, default=0,
                        help="源数据库端口号（0=使用默认端口, MySQL:3306, PostgreSQL:5432）")
    s_src.add_argument("--source-user", default="", help="源数据库用户名")
    s_src.add_argument("--source-password", default="", help="源数据库密码")
    s_src.add_argument("--source-database", default="", help="源数据库名 (MySQL/PostgreSQL)")
    s_src.add_argument("--source-path", default="", help="SQLite 数据库文件路径")

    # 目标数据库
    s_tgt = p_migrate.add_argument_group("目标数据库配置")
    s_tgt.add_argument("--target-type", required=True,
                        choices=["sqlite", "mysql", "postgresql", "pg", "psql"],
                        help="目标数据库类型")
    s_tgt.add_argument("--target-host", default="127.0.0.1",
                        help="目标数据库主机地址 (默认 127.0.0.1)")
    s_tgt.add_argument("--target-port", type=int, default=0,
                        help="目标数据库端口号")
    s_tgt.add_argument("--target-user", default="", help="目标数据库用户名")
    s_tgt.add_argument("--target-password", default="", help="目标数据库密码")
    s_tgt.add_argument("--target-database", default="", help="目标数据库名")
    s_tgt.add_argument("--target-path", default="", help="SQLite 目标文件路径")

    # 迁移控制
    s_ctl = p_migrate.add_argument_group("迁移控制")
    s_ctl.add_argument("--tables", default=None,
                        help="指定要迁移的表名（逗号分隔，如: users,orders），不指定则迁移全部表")
    s_ctl.add_argument("--threads", type=int, default=8,
                        help="最大线程数（默认 8，最多 20，数据超过 5 万行时生效）")
    s_ctl.add_argument("--batch-size", type=int, default=2500,
                        help="分批处理条数（默认 2500，数据超过 3000 行时生效）")
    s_ctl.add_argument("--drop-if-exists", action="store_true",
                        help="若目标表已存在，先删除再重建（危险操作，默认关闭）")
    s_ctl.add_argument("--create-tables", action="store_true", default=True,
                        help="在目标数据库自动创建表结构（默认开启）")
    s_ctl.add_argument("-y", "--yes", action="store_true", default=False,
                        help="跳过迁移前的二次确认")
    s_ctl.add_argument("-v", "--verbose", action="store_true",
                        help="输出更多调试信息")

    # ---------------- export 子命令 ----------------
    p_export = subparsers.add_parser(
        "export",
        help="将数据库表导出为 Excel(.xlsx) 或 CSV",
        description="数据库导出 - 支持 SQLite / MySQL / PostgreSQL 导出为 Excel 或 CSV，"
                    "超过 3000 行自动分批（每批 2500 条），单 sheet 最多 5 万行，"
                    "单工作簿 50 万行自动分文件，多文件时自动创建子文件夹。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    e_src = p_export.add_argument_group("源数据库配置")
    e_src.add_argument("--source-type", required=True,
                        choices=["sqlite", "mysql", "postgresql", "pg", "psql"],
                        help="源数据库类型")
    e_src.add_argument("--source-host", default="127.0.0.1",
                        help="源数据库主机地址 (默认 127.0.0.1)")
    e_src.add_argument("--source-port", type=int, default=0, help="源数据库端口号")
    e_src.add_argument("--source-user", default="", help="源数据库用户名")
    e_src.add_argument("--source-password", default="", help="源数据库密码")
    e_src.add_argument("--source-database", default="", help="源数据库名 (MySQL/PostgreSQL)")
    e_src.add_argument("--source-path", default="", help="SQLite 数据库文件路径")

    e_ctl = p_export.add_argument_group("导出控制")
    e_ctl.add_argument("--tables", default=None,
                        help="指定要导出的表名（逗号分隔，如: users,orders），不指定则导出全部表")
    e_ctl.add_argument("--format", "--output-format", dest="output_format",
                        choices=["xlsx", "csv"], default="xlsx",
                        help="导出格式，默认 xlsx（Excel）")
    e_ctl.add_argument("--output-dir", default="./exports",
                        help="输出目录（默认 ./exports）")
    e_ctl.add_argument("--filename", default=None,
                        help="自定义文件名（不含扩展名）；留空则自动生成: 数据库_表名_行数_年月日_时分秒")
    e_ctl.add_argument("--from-row", type=int, default=0,
                        help="从第几行开始导出（从 0 开始计数，默认 0）")
    e_ctl.add_argument("--to-row", type=int, default=None,
                        help="导出到第几行（不含该行，默认到最后）")
    e_ctl.add_argument("--limit", type=int, default=None,
                        help="最多导出多少行（优先级低于 --to-row）")
    e_ctl.add_argument("--threads", type=int, default=8,
                        help="最大线程数（默认 8，最多 20，MySQL/PostgreSQL 超过 50 万行时启用多线程，"
                             "SQLite 文件锁不支持多线程加速，会自动回退到大批次单线程）")
    e_ctl.add_argument("--batch-size", type=int, default=2500,
                        help="分批读取条数（默认 2500，超过 3000 行时生效）")
    e_ctl.add_argument("-y", "--yes", action="store_true", default=False,
                        help="跳过导出前的二次确认")

    return parser


def main() -> int:
    parser = _build_parser()

    # 无参数或仅有 -h：显示主帮助
    if len(sys.argv) < 2:
        parser.print_help()
        return 0
    first_arg = sys.argv[1]

    # 单独 -h / --help / help
    if first_arg in ("-h", "--help", "help") and len(sys.argv) == 2:
        parser.print_help()
        return 0

    # 识别子命令
    known_subcommands = {"migrate", "migration", "export", "dump"}
    subcommand = None
    if first_arg in known_subcommands:
        subcommand = "migrate" if first_arg in ("migrate", "migration") else "export"
        argv_for_sub = [sys.argv[0]] + sys.argv[2:]
    else:
        # 未识别子命令 → 默认为 migrate（向后兼容老用法）
        subcommand = "migrate"
        argv_for_sub = sys.argv[:]

    # 根据子命令交给对应模块解析和执行
    if subcommand == "migrate":
        sys.argv = argv_for_sub
        from core.db_migration_tool import main as migrate_main
        return migrate_main()

    sys.argv = argv_for_sub
    from core.db_export_cli import main as export_main
    return export_main()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n[INFO] 用户中断操作")
        sys.exit(130)
# 数据库迁移与导出工具 | Database Migration & Export Tool

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Database](https://img.shields.io/badge/db-SQLite%20%7C%20MySQL%20%7C%20PostgreSQL-orange.svg)]()

> **语言切换 | Language Switch**: [中文（当前）](#) | [English](#english-version)

支持 **SQLite / MySQL / PostgreSQL** 三种数据库之间的互迁移，以及将数据库表导出为 **Excel(.xlsx) / CSV** 文件。

*A Python tool for cross-database migration (SQLite ↔ MySQL ↔ PostgreSQL) and exporting tables to Excel(.xlsx) / CSV.*

---

## 目录结构 | Project Structure

```
sql_migrate_tools/
├── run.py                   ← 启动脚本（统一入口，支持子命令）
├── requirements.txt         ← Python 依赖清单
├── README.md                ← 本文件
├── core/                    ← 核心模块（内部包）
│   ├── __init__.py
│   ├── database_connector.py   ← 数据库连接管理（SQLite / MySQL / PG）
│   ├── database_migrator.py    ← 迁移核心（分批 / 多线程 / 事务 / 主键分页）
│   ├── db_migration_tool.py    ← migrate 子命令参数解析与执行
│   ├── db_exporter.py          ← 导出引擎（分批 / 多线程 / 多进程 / 分 Sheet / 分文件）
│   ├── db_export_cli.py        ← export 子命令参数解析与执行
│   ├── sql_dialect.py          ← SQL 方言转换（25+ 种数据类型映射）
│   └── progress_bar.py         ← 统一终端进度条（ETA 预估 / 多线程安全）
└── test/                   ← 测试脚本
    ├── create_db2.py
    └── generate_test_db.py
```

---

## 功能特性

### 数据库迁移（migrate）

| 功能 | 说明 |
|------|------|
| 三库互迁移 | SQLite ↔ MySQL ↔ PostgreSQL 任意组合，支持 9 种迁移方向 |
| 智能分页策略 | 自动检测主键类型，优先使用整数主键范围查询（O(1)），其次键集分页（O(1)），最后 LIMIT/OFFSET 回退 |
| 分批处理 | 单表数据超过 **3000 条** 时自动启用，默认每批 **10000 条**（可自定义） |
| 多线程写入 | 单表数据超过 **50000 条** 且非 SQLite 源时自动启用，最多 **20 个线程** |
| PostgreSQL 优化 | 使用 `execute_values` 批量写入，性能提升 **10-30 倍** |
| MySQL 自增主键 | 迁移到 MySQL 时，单列整数主键自动添加 `AUTO_INCREMENT` |
| SQLite 保护 | 检测到 SQLite 源时自动回退到单线程模式，避免文件锁竞争 |
| 事务保障 | 迁移过程中任意表失败即 **全部回滚**，不会产生脏数据 |
| 表名映射 | 通过 `--rename` 参数支持源表到目标表的重命名映射 |
| 灵活指定 | 可通过 `--tables` 指定表，或 `--databases` 指定库 |
| 驱动检测 | 启动时自动检查目标数据库的 Python 驱动是否安装 |
| 中文日志 | 完整的中文进度与错误日志，带时间戳和级别标记 |

### 数据库导出（export）

| 功能 | 说明 |
|------|------|
| 多格式导出 | 支持 Excel(.xlsx) 和 CSV 两种格式 |
| 双引擎 Excel | 优先使用 xlsxwriter（更快），回退 openpyxl |
| 分批读取 | 数据超过 **3000 条** 自动分批，默认每批 **50000 条**（可自定义） |
| MySQL 流式读取 | 使用 SSCursor（服务端游标）流式读取，避免全量加载到内存 |
| 多线程读取 | 数据超过 **100000 行**（MySQL/PG）时启用多线程并行读取 |
| 多进程分段 | 数据超过 **500000 行** 时启用多进程分段导出（最多 4 进程） |
| Sheet 自动切分 | 单 Sheet 最多 **50000 行**，超出自动创建新 Sheet |
| 工作簿自动切分 | 单工作簿最多 **500000 行**，超出自动创建新文件（如 `_w002.xlsx`） |
| 行范围导出 | 通过 `--from-row` 和 `--to-row` 指定导出范围 |
| 限制行数 | 通过 `--limit` 限制最大导出行数 |
| 自定义文件名 | 支持 `--filename` 自定义，留空则自动生成（含数据库名、表名、行数、时间戳） |
| 文件完整性校验 | 导出完成后自动校验 ZIP 结构（xlsx）和文件大小 |
| 智能分页 | 同迁移模块，自动检测主键并选择最优分页策略 |

---

## 依赖安装

```bash
# SQLite: Python 标准库自带，无需额外安装
# MySQL: 需要 pymysql
pip install pymysql

# PostgreSQL: 需要 psycopg2-binary
pip install psycopg2-binary

# Excel 导出: 需要 xlsxwriter 或 openpyxl
pip install xlsxwriter

# 或一次性安装全部
pip install -r requirements.txt
```

---

## 快速开始

```bash
# 查看总帮助
python run.py -h

# 查看 migrate 子命令帮助
python run.py migrate -h

# 查看 export 子命令帮助
python run.py export -h
```

---

## 子命令一：数据库迁移（migrate）

### 使用示例

#### 示例 1：SQLite → MySQL（迁移全部表）

```bash
python run.py migrate \
    --source-type sqlite --source-path ./data/source.db \
    --target-type mysql \
    --target-host 127.0.0.1 \
    --target-port 3306 \
    --target-user root \
    --target-password 123456 \
    --target-database target_db \
    -y
```

- 迁移 `source.db` 中 **所有表** 到 MySQL 的 `target_db` 库
- `-y` 跳过迁移前确认
- 目标数据库 `target_db` 需要 **预先创建**（MySQL/PG）

#### 示例 2：MySQL → PostgreSQL（指定表，多线程）

```bash
python run.py migrate \
    --source-type mysql \
    --source-host 127.0.0.1 \
    --source-port 3306 \
    --source-user root \
    --source-password 123456 \
    --source-database src_db \
    --target-type postgresql \
    --target-host 127.0.0.1 \
    --target-port 5432 \
    --target-user postgres \
    --target-password 123456 \
    --target-database target_db \
    --tables users,orders,products \
    --threads 10 \
    --batch-size 5000 \
    -y
```

- 只迁移 `users`、`orders`、`products` 三张表
- `--threads 10`：大数据量时使用 10 个并发线程
- `--batch-size 5000`：每批处理 5000 条（默认 10000）

#### 示例 3：PostgreSQL → SQLite（单表迁移）

```bash
python run.py migrate \
    --source-type postgresql \
    --source-host 127.0.0.1 \
    --source-port 5432 \
    --source-user postgres \
    --source-password 123456 \
    --source-database src_db \
    --target-type sqlite --target-path ./backup/export.db \
    --tables employees \
    --batch-size 5000 \
    -y
```

- 将 PG 的 `employees` 表迁移到 SQLite 文件
- 小批量处理适合网络不稳定的场景
- `./backup/` 目录不存在时会自动创建

#### 示例 4：MySQL → MySQL（跨库迁移 + 表名映射）

```bash
python run.py migrate \
    --source-type mysql --source-host 127.0.0.1 --source-user root \
    --source-password 123456 --source-database old_db \
    --target-type mysql --target-host 192.168.1.100 --target-user root \
    --target-password 123456 --target-database new_db \
    --rename users:t_users,orders:t_orders \
    -y
```

- `--rename` 将源表 `users` 映射为目标表 `t_users`，`orders` 映射为 `t_orders`

#### 示例 5：SQLite → SQLite（本地测试验证）

```bash
python run.py migrate \
    --source-type sqlite --source-path ./data/demo.db \
    --target-type sqlite --target-path ./data/copy.db -y
```

这是最简单的测试方式，不需要安装任何数据库驱动。

---

### migrate 参数详解

#### 源数据库参数（`--source-*`）

| 参数 | 必填 | 适用数据库 | 说明 |
|------|------|----------|------|
| `--source-type` | ✅ | 全部 | `sqlite` / `mysql` / `postgresql` / `psql` / `pg` |
| `--source-host` | - | MySQL / PG | 主机地址，默认 `127.0.0.1` |
| `--source-port` | - | MySQL / PG | 端口，不填则使用数据库默认端口（MySQL:3306, PG:5432） |
| `--source-user` | - | MySQL / PG | 用户名 |
| `--source-password` | - | MySQL / PG | 密码 |
| `--source-database` | - | MySQL / PG | 数据库名 |
| `--source-path` | - | SQLite | 数据库文件路径（如 `./data/src.db`） |
| `--source-charset` | - | MySQL | 字符集，默认 `utf8mb4` |

#### 目标数据库参数（`--target-*`）

同上，把前缀 `source` 换成 `target` 即可。

#### 迁移控制参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--tables` | - | 指定要迁移的表名（逗号分隔），如 `users,orders,products`。**不填则迁移全部表** |
| `--rename` | - | 表名映射（源表:目标表，逗号分隔多对），如 `users:t_users,orders:t_orders` |
| `--databases` | - | 指定要迁移的数据库（仅支持 MySQL 多库场景） |
| `--threads` | 8 | 最大线程数（仅单表数据超过 5 万条时生效），最多 20 |
| `--batch-size` | 10000 | 每批处理条数（仅单表数据超过 3000 条时生效） |
| `-y` / `--yes` | False | 跳过迁移前的二次确认 |
| `-v` / `--verbose` | False | 输出更多调试信息 |

---

## 子命令二：数据库导出（export）

### 使用示例

#### 示例 1：导出 SQLite 全部表为 Excel

```bash
python run.py export \
    --source-type sqlite --source-path ./data/source.db \
    --format xlsx --output-dir ./exports -y
```

#### 示例 2：导出 MySQL 指定表为 CSV

```bash
python run.py export \
    --source-type mysql --source-host 127.0.0.1 \
    --source-port 3306 --source-user root --source-password 123456 \
    --source-database test_db --tables users,orders \
    --format csv --filename my_export -y
```

#### 示例 3：导出 PostgreSQL 指定行范围

```bash
python run.py export \
    --source-type postgresql --source-host 127.0.0.1 \
    --source-port 5432 --source-user postgres --source-password 123456 \
    --source-database test_db --tables products \
    --from-row 100 --to-row 600 -y
```

- 导出 `products` 表的第 100 到 600 行

#### 示例 4：导出前 500 行

```bash
python run.py export \
    --source-type sqlite --source-path ./data/source.db \
    --tables employees --limit 500 -y
```

#### 示例 5：超大表导出（自定义分批）

```bash
python run.py export \
    --source-type mysql --source-host 127.0.0.1 \
    --source-user root --source-password 123456 --source-database big_db \
    --tables huge_table --batch-size 50000 --threads 8 -y
```

---

### export 参数详解

#### 源数据库参数（`--source-*`）

| 参数 | 必填 | 适用数据库 | 说明 |
|------|------|----------|------|
| `--source-type` | ✅ | 全部 | `sqlite` / `mysql` / `postgresql` / `psql` / `pg` |
| `--source-host` | - | MySQL / PG | 主机地址，默认 `127.0.0.1` |
| `--source-port` | - | MySQL / PG | 端口，不填则使用默认端口 |
| `--source-user` | - | MySQL / PG | 用户名 |
| `--source-password` | - | MySQL / PG | 密码 |
| `--source-database` | - | MySQL / PG | 数据库名 |
| `--source-path` | - | SQLite | 数据库文件路径 |
| `--source-charset` | - | MySQL | 字符集，默认 `utf8mb4` |

#### 导出控制参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--tables` | - | 指定要导出的表名（逗号分隔），不填则导出全部表 |
| `--format` | xlsx | 导出格式，可选 `xlsx`（Excel）或 `csv` |
| `--output-dir` | ./exports | 输出目录 |
| `--filename` | - | 自定义文件名（不含扩展名）。留空则自动生成: `数据库名_表名_行数_年月日_时分秒` |
| `--from-row` | 0 | 从第几行开始导出（从 0 开始计数） |
| `--to-row` | - | 导出到第几行（不含该行） |
| `--limit` | - | 最多导出多少行 |
| `--threads` | 8 | 最大线程数（默认 8，最多 20，MySQL/PostgreSQL 超过 10 万行时启用多线程） |
| `--batch-size` | 50000 | 分批读取条数（默认 50000，超过 3000 行时启用分批） |
| `-y` / `--yes` | False | 跳过导出前的二次确认 |

---

## 重要说明

### 1. 目标数据库需预先创建

- **SQLite**：文件路径无需预先创建，工具会自动创建目录和文件
- **MySQL / PostgreSQL**：目标数据库（database）必须 **预先手工创建**
  - MySQL: `CREATE DATABASE target_db DEFAULT CHARACTER SET utf8mb4;`
  - PG: `CREATE DATABASE target_db;`

### 2. 迁移行为

- 目标库中的 **同名表会被删除并重建**（DROP TABLE IF EXISTS → CREATE TABLE）
- 仅迁移表结构和数据，**不迁移索引、外键、触发器、存储过程** 等复杂对象
- 数据类型会自动转换为目标数据库兼容的类型（25+ 种类型映射）
- 可使用 `--rename` 将源表映射到不同名称的目标表

### 3. 分批与多线程的触发条件

| 单表数据量 | 处理方式 |
|----------|---------|
| ≤ 3000 条 | 一次性插入/读取 |
| 3001 ~ 50000 条 | 单线程分批（每批 10000 条） |
| > 50000 条 | 多线程分批（线程数 = min(threads, 批次数)） |

### 4. 智能分页策略（性能核心）

工具会自动检测主键类型并选择最优分页策略：

| 策略 | 触发条件 | 复杂度 | 性能 |
|------|---------|--------|------|
| 整数主键范围查询 | 单列整数主键 | O(1) | 最快（270 万行 < 2 分钟） |
| 键集分页 | 单列字符串主键 | O(1) 每批 | 快（比 LIMIT/OFFSET 快 10-50 倍） |
| LIMIT/OFFSET | 无主键 | O(n²) | 慢（大表建议添加主键） |

### 5. 事务与回滚

- 任意一张表迁移失败 → 整个迁移任务终止 → 所有已写入的数据 **全部回滚**
- 错误发生时会打印失败的表名和具体错误信息

### 6. 导出大表时的自动切分

| 阈值 | 行为 |
|------|------|
| 50000 行/Sheet | 自动创建新 Sheet（如 `sheet0`, `sheet1`...） |
| 500000 行/工作簿 | 自动创建新文件（如 `xxx_w002.xlsx`, `xxx_w003.xlsx`...） |
| 500000 行/表 | 启用多进程分段导出（最多 4 进程） |

---

## 退出码说明

| 退出码 | 含义 |
|--------|------|
| 0 | 成功（或用户主动取消） |
| 1 | 缺少数据库驱动 |
| 2 | 参数配置错误 |
| 3 | 迁移模块加载失败 |
| 4 | 数据库连接失败 |
| 5 | 迁移/导出过程失败（迁移已回滚） |
| 130 | 用户按 Ctrl+C 中断 |
| 99 | 未预期的运行时错误 |

---

## 日志示例（正常迁移）

```
============================================================
正在检查 Python 数据库驱动依赖...
[INFO] 数据库驱动 sqlite3 已安装 ✓
[INFO] 数据库驱动 sqlite3 已安装 ✓
============================================================

============================================================
【迁移计划摘要】
------------------------------------------------------------
源数据库: SQLITE
  路径: ./data/source.db
目标数据库: SQLITE
  路径: ./data/backup.db
指定表: employees, orders（共 2 张）
最大线程数: 4
分批大小: 500
============================================================

[INFO] 2026-07-18 15:00:00 正在测试源数据库连接...
[SUCCESS] 2026-07-18 15:00:00 源数据库连接正常
[INFO] 2026-07-18 15:00:00 正在测试目标数据库连接...
[SUCCESS] 2026-07-18 15:00:00 目标数据库连接正常
[INFO] 2026-07-18 15:00:00 源数据库: sqlite -> 目标数据库: sqlite
[INFO] 2026-07-18 15:00:00 迁移计划共 2 张表
[INFO] 2026-07-18 15:00:00 开始迁移表 employees，共 1500 条记录
[INFO] 2026-07-18 15:00:00 表 employees 已在目标库创建/重建
[INFO] 2026-07-18 15:00:00 数据量 1500 超过 3000 条，采用分批处理（每批 500 条，共 3 批）
[INFO] 2026-07-18 15:00:01 表 employees 进度: 3/3 批
[SUCCESS] 2026-07-18 15:00:01 表 employees 迁移完成，耗时 0.05s，共处理 1500 条记录
[SUCCESS] 2026-07-18 15:00:01 全部迁移完成！共处理 2 张表，1700 条记录，总耗时 0.06s

============================================================
【迁移结果汇总】
------------------------------------------------------------
  ✓ employees: 1500 行，3 批，1 线程，耗时 0.05s
  ✓ orders: 200 行，1 批，1 线程，耗时 0.00s
------------------------------------------------------------
总计: 2 张表，1700 条记录
总耗时: 0.06 秒
============================================================
```

---

## 数据类型映射表

工具内置 25+ 种数据类型的自动映射，以下为部分示例：

| 原始类型 | SQLite | MySQL | PostgreSQL |
|---------|--------|-------|------------|
| INTEGER / INT | INTEGER | INT | INTEGER |
| BIGINT | INTEGER | BIGINT | BIGINT |
| VARCHAR | TEXT | VARCHAR(255) | VARCHAR |
| TEXT | TEXT | TEXT | TEXT |
| JSON | TEXT | JSON | JSONB |
| BOOLEAN | INTEGER | TINYINT(1) | BOOLEAN |
| DATETIME | TEXT | DATETIME | TIMESTAMP |
| FLOAT | REAL | FLOAT | REAL |
| DOUBLE | REAL | DOUBLE | DOUBLE PRECISION |
| DECIMAL | REAL | DECIMAL(20,6) | DECIMAL(20,6) |
| BLOB | BLOB | BLOB | BYTEA |
| UUID | TEXT | VARCHAR(36) | UUID |

---

## 常见问题

**Q: 迁移报错 "缺少数据库驱动"？**
A: 安装对应驱动：`pip install pymysql`（MySQL）、`pip install psycopg2-binary`（PostgreSQL）、`pip install xlsxwriter`（Excel 导出）。

**Q: MySQL 迁移失败，提示 "Access denied"？**
A: 检查用户名、密码、端口是否正确，以及用户是否有目标数据库的 CREATE TABLE 和 INSERT 权限。

**Q: 迁移大表时内存占用会不会很高？**
A: 不会，超过 3000 条会自动分批读取和写入，内存占用稳定在很低的水平。

**Q: 为什么 PostgreSQL 看不到我要迁移的表？**
A: 工具默认只读取 `public` schema 下的表。如需迁移其他 schema，请在 --source-database 指定对应 schema。

**Q: 为什么目标表中的数据和源表一致，但顺序不同？**
A: 迁移时不保证数据顺序（因为 SQL 本身不保证顺序），但数据内容和条数完全一致。

**Q: 导出 Excel 报错 "需要安装 xlsxwriter 或 openpyxl"？**
A: 安装任一引擎即可：`pip install xlsxwriter`（推荐，更快）或 `pip install openpyxl`。

**Q: 如何将源表迁移到不同名称的目标表？**
A: 使用 `--rename` 参数：`--rename users:t_users,orders:t_orders`。

**Q: 导出超大表（千万级）时会不会内存溢出？**
A: 不会。工具采用流式读取 + 分批写入，同时自动切分 Sheet（5 万行/Sheet）和工作簿（50 万行/工作簿），确保内存始终可控。

---

## 快速测试

```bash
# 1. 先查看帮助确认参数
python run.py -h

# 2. SQLite -> SQLite 测试（不需要数据库服务器）
python -c "import sqlite3;c=sqlite3.connect('test.db');[c.execute('CREATE TABLE t'+str(i)+'(id INTEGER PRIMARY KEY, name TEXT)') for i in range(3)];[c.execute('INSERT INTO t'+str(i)+'(name) VALUES(\"name_\")') for i in range(3) for _ in range(100)];c.commit();c.close();print('test.db 已创建')"
python run.py migrate --source-type sqlite --source-path test.db --target-type sqlite --target-path test_backup.db -y

# 3. 导出测试
python run.py export --source-type sqlite --source-path test.db --format csv -y
```

验证结果：
```bash
python -c "import sqlite3;s=sqlite3.connect('test.db');d=sqlite3.connect('test_backup.db');[print(f't{i}: src={s.execute(\"SELECT COUNT(*) FROM t\"+str(i)).fetchone()[0]}, dst={d.execute(\"SELECT COUNT(*) FROM t\"+str(i)).fetchone()[0]}') for i in range(3)]"
```

---

## 免责声明 | Disclaimer

> **中文**：本工具按"原样"提供，不提供任何形式的明示或暗示担保。使用本工具进行数据库迁移或导出操作前，请务必先备份您的数据。作者不对因使用本工具而导致的任何数据丢失、损坏或业务中断承担责任。在生产环境使用前，请先在测试环境中充分验证。
>
> **English**: This tool is provided "AS IS", without warranty of any kind, express or implied. Always back up your data before performing database migration or export operations with this tool. The author shall not be liable for any data loss, corruption, or business interruption caused by the use of this tool. Please thoroughly test in a staging environment before using in production.

---

## 协议 | License

本项目基于 **MIT License** 开源协议发布。详见 [LICENSE](LICENSE) 文件。

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

## English Version

> **[中文版本](#数据库迁移与导出工具--database-migration--export-tool)**

### Project Structure

```
sql_migrate_tools/
├── run.py                       ← Entry point (subcommands: migrate / export)
├── requirements.txt             ← Python dependencies
├── README.md                    ← This file
├── LICENSE                      ← MIT License
├── core/                        ← Core modules
│   ├── __init__.py
│   ├── database_connector.py    ← DB connection management (SQLite / MySQL / PG)
│   ├── database_migrator.py     ← Migration engine (batching / threading / transactions / PK pagination)
│   ├── db_migration_tool.py     ← migrate subcommand CLI
│   ├── db_exporter.py           ← Export engine (batching / threading / multi-process / sheet split / file split)
│   ├── db_export_cli.py         ← export subcommand CLI
│   ├── sql_dialect.py           ← SQL dialect converter (25+ data type mappings)
│   └── progress_bar.py          ← Unified progress bar (ETA estimation / thread-safe)
└── test/                        ← Test scripts
    ├── create_db2.py
    └── generate_test_db.py
```

### Overview

A powerful Python tool for cross-database migration between **SQLite**, **MySQL**, and **PostgreSQL**, and exporting database tables to **Excel (.xlsx)** or **CSV** files.

### Features

#### Database Migration

| Feature | Description |
|---------|-------------|
| Cross-DB Migration | Any combination of SQLite ↔ MySQL ↔ PostgreSQL (9 directions) |
| Smart Pagination | Auto-detects primary key type, prefers integer range query (O(1)), then keyset pagination (O(1)), falls back to LIMIT/OFFSET |
| Batch Processing | Auto-enabled for tables with **3,000+** rows, default **10,000** rows/batch |
| Multi-threading | Auto-enabled for tables with **50,000+** rows (non-SQLite source), up to **20 threads** |
| PostgreSQL Optimization | Uses `execute_values` for batch inserts, **10-30x** performance boost |
| MySQL AUTO_INCREMENT | Auto-adds `AUTO_INCREMENT` for single-column integer primary keys |
| SQLite Protection | Auto-falls back to single-threaded mode to avoid file lock contention |
| Transaction Safety | Any table failure → full rollback, no dirty data |
| Table Rename Mapping | `--rename` parameter for source-to-target table name mapping |
| Driver Detection | Auto-checks Python DB drivers on startup |

#### Database Export

| Feature | Description |
|---------|-------------|
| Multi-format | Supports Excel (.xlsx) and CSV |
| Dual Excel Engines | Prefers xlsxwriter (faster), falls back to openpyxl |
| Batch Reading | Auto-batches for **3,000+** rows, default **50,000** rows/batch |
| MySQL Streaming | Uses SSCursor (server-side cursor) for streaming reads |
| Multi-threaded Read | Enabled for **100,000+** rows (MySQL/PG) |
| Multi-process Export | Enabled for **500,000+** rows per table (up to 4 processes) |
| Sheet Auto-split | Max **50,000** rows per sheet, auto-creates new sheets |
| Workbook Auto-split | Max **500,000** rows per workbook, auto-creates new files |
| Row Range Export | `--from-row` and `--to-row` for range selection |
| Row Limit | `--limit` parameter for capping rows |
| Custom Filename | `--filename` support, auto-generates with DB name, table name, row count, timestamp |
| File Integrity Check | Auto-validates ZIP structure (xlsx) and file size post-export |

### Installation

```bash
# SQLite: built into Python stdlib, no extra install needed
# MySQL: requires pymysql
pip install pymysql

# PostgreSQL: requires psycopg2-binary
pip install psycopg2-binary

# Excel export: requires xlsxwriter or openpyxl
pip install xlsxwriter

# Or install all at once
pip install -r requirements.txt
```

### Quick Start

```bash
# View all help
python run.py -h

# View migrate subcommand help
python run.py migrate -h

# View export subcommand help
python run.py export -h
```

---

### Subcommand 1: Database Migration (`migrate`)

#### Examples

**Example 1: SQLite → MySQL (all tables)**

```bash
python run.py migrate \
    --source-type sqlite --source-path ./data/source.db \
    --target-type mysql \
    --target-host 127.0.0.1 \
    --target-port 3306 \
    --target-user root \
    --target-password 123456 \
    --target-database target_db \
    -y
```

- Migrates **all tables** from `source.db` to MySQL's `target_db`
- `-y` skips the confirmation prompt
- The target database `target_db` must be **created beforehand** (MySQL/PG)

**Example 2: MySQL → PostgreSQL (specific tables, multi-threaded)**

```bash
python run.py migrate \
    --source-type mysql \
    --source-host 127.0.0.1 \
    --source-port 3306 \
    --source-user root \
    --source-password 123456 \
    --source-database src_db \
    --target-type postgresql \
    --target-host 127.0.0.1 \
    --target-port 5432 \
    --target-user postgres \
    --target-password 123456 \
    --target-database target_db \
    --tables users,orders,products \
    --threads 10 \
    --batch-size 5000 \
    -y
```

- Only migrates `users`, `orders`, `products` tables
- `--threads 10`: uses 10 concurrent threads for large tables
- `--batch-size 5000`: processes 5000 rows per batch (default 10000)

**Example 3: PostgreSQL → SQLite (single table)**

```bash
python run.py migrate \
    --source-type postgresql \
    --source-host 127.0.0.1 \
    --source-port 5432 \
    --source-user postgres \
    --source-password 123456 \
    --source-database src_db \
    --target-type sqlite --target-path ./backup/export.db \
    --tables employees \
    --batch-size 5000 \
    -y
```

- Migrates the `employees` table from PG to a SQLite file
- The `./backup/` directory is auto-created if it doesn't exist

**Example 4: MySQL → MySQL (cross-server migration + table rename)**

```bash
python run.py migrate \
    --source-type mysql --source-host 127.0.0.1 --source-user root \
    --source-password 123456 --source-database old_db \
    --target-type mysql --target-host 192.168.1.100 --target-user root \
    --target-password 123456 --target-database new_db \
    --rename users:t_users,orders:t_orders \
    -y
```

- `--rename` maps source `users` → target `t_users`, `orders` → `t_orders`

**Example 5: SQLite → SQLite (local test)**

```bash
python run.py migrate \
    --source-type sqlite --source-path ./data/demo.db \
    --target-type sqlite --target-path ./data/copy.db -y
```

The simplest test scenario — no database server required.

#### `migrate` Parameters

**Source Database (`--source-*`)**

| Parameter | Required | Applies To | Description |
|-----------|----------|------------|-------------|
| `--source-type` | ✅ | All | `sqlite` / `mysql` / `postgresql` / `psql` / `pg` |
| `--source-host` | - | MySQL / PG | Host address, default `127.0.0.1` |
| `--source-port` | - | MySQL / PG | Port, defaults to DB default (MySQL:3306, PG:5432) |
| `--source-user` | - | MySQL / PG | Username |
| `--source-password` | - | MySQL / PG | Password |
| `--source-database` | - | MySQL / PG | Database name |
| `--source-path` | - | SQLite | Database file path (e.g., `./data/src.db`) |
| `--source-charset` | - | MySQL | Character set, default `utf8mb4` |

**Target Database (`--target-*`)**

Same as above, replace `source` with `target`.

**Migration Control**

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--tables` | - | Tables to migrate (comma-separated). **Omit to migrate all** |
| `--rename` | - | Table name mapping (`src:dst`, comma-separated pairs), e.g. `users:t_users,orders:t_orders` |
| `--databases` | - | Databases to migrate (MySQL multi-DB scenario only) |
| `--threads` | 8 | Max threads (only for tables with 50,000+ rows), max 20 |
| `--batch-size` | 10000 | Rows per batch (only for tables with 3,000+ rows) |
| `-y` / `--yes` | False | Skip confirmation prompt |
| `-v` / `--verbose` | False | Verbose debug output |

---

### Subcommand 2: Database Export (`export`)

#### Examples

**Example 1: Export all SQLite tables to Excel**

```bash
python run.py export \
    --source-type sqlite --source-path ./data/source.db \
    --format xlsx --output-dir ./exports -y
```

**Example 2: Export MySQL tables to CSV with custom filename**

```bash
python run.py export \
    --source-type mysql --source-host 127.0.0.1 \
    --source-port 3306 --source-user root --source-password 123456 \
    --source-database test_db --tables users,orders \
    --format csv --filename my_export -y
```

**Example 3: Export PostgreSQL with row range**

```bash
python run.py export \
    --source-type postgresql --source-host 127.0.0.1 \
    --source-port 5432 --source-user postgres --source-password 123456 \
    --source-database test_db --tables products \
    --from-row 100 --to-row 600 -y
```

- Exports rows 100–599 from the `products` table

**Example 4: Export first 500 rows**

```bash
python run.py export \
    --source-type sqlite --source-path ./data/source.db \
    --tables employees --limit 500 -y
```

**Example 5: Export large table with custom batch size**

```bash
python run.py export \
    --source-type mysql --source-host 127.0.0.1 \
    --source-user root --source-password 123456 --source-database big_db \
    --tables huge_table --batch-size 50000 --threads 8 -y
```

#### `export` Parameters

**Source Database (`--source-*`)**

| Parameter | Required | Applies To | Description |
|-----------|----------|------------|-------------|
| `--source-type` | ✅ | All | `sqlite` / `mysql` / `postgresql` / `psql` / `pg` |
| `--source-host` | - | MySQL / PG | Host address, default `127.0.0.1` |
| `--source-port` | - | MySQL / PG | Port, defaults to DB default |
| `--source-user` | - | MySQL / PG | Username |
| `--source-password` | - | MySQL / PG | Password |
| `--source-database` | - | MySQL / PG | Database name |
| `--source-path` | - | SQLite | Database file path |
| `--source-charset` | - | MySQL | Character set, default `utf8mb4` |

**Export Control**

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--tables` | - | Tables to export (comma-separated). **Omit to export all** |
| `--format` | xlsx | Export format: `xlsx` (Excel) or `csv` |
| `--output-dir` | ./exports | Output directory |
| `--filename` | - | Custom filename (no extension). Auto-generates with DB name, table name, row count, timestamp if omitted |
| `--from-row` | 0 | Start exporting from this row (0-indexed) |
| `--to-row` | - | Export up to (but not including) this row |
| `--limit` | - | Maximum number of rows to export |
| `--threads` | 8 | Max threads (default 8, max 20, multi-threading enabled for MySQL/PG tables with 100,000+ rows) |
| `--batch-size` | 50000 | Batch read size (default 50000, batching enabled for 3,000+ rows) |
| `-y` / `--yes` | False | Skip confirmation prompt |

---

### Important Notes

#### 1. Target Database Must Be Pre-created

- **SQLite**: file path is auto-created; no manual setup needed
- **MySQL / PostgreSQL**: the target database must be **created manually** beforehand
  - MySQL: `CREATE DATABASE target_db DEFAULT CHARACTER SET utf8mb4;`
  - PG: `CREATE DATABASE target_db;`

#### 2. Migration Behavior

- **Existing tables in the target are dropped and recreated** (DROP TABLE IF EXISTS → CREATE TABLE)
- Only table structure and data are migrated; **indexes, foreign keys, triggers, and stored procedures are NOT migrated**
- Data types are automatically converted to compatible types for the target database (25+ type mappings)
- Use `--rename` to map source tables to differently-named target tables

#### 3. Batch & Multi-threading Trigger Conditions

| Table Row Count | Processing Mode |
|-----------------|-----------------|
| ≤ 3,000 | Direct insert/read (single batch) |
| 3,001 ~ 50,000 | Single-threaded batching (10,000 rows/batch) |
| > 50,000 | Multi-threaded batching (threads = min(threads, batch_count)) |

#### 4. Smart Pagination Strategy (Performance Core)

The tool automatically detects the primary key type and selects the optimal pagination strategy:

| Strategy | Trigger Condition | Complexity | Performance |
|----------|-------------------|------------|-------------|
| Integer PK Range Query | Single-column integer PK | O(1) | Fastest (2.7M rows < 2 min) |
| Keyset Pagination | Single-column string PK | O(1) per batch | Fast (10-50x faster than LIMIT/OFFSET) |
| LIMIT/OFFSET | No PK | O(n²) | Slow (add a PK for better performance) |

#### 5. Transactions & Rollback

- If **any** table fails during migration → the entire task is terminated → **all data already written is rolled back**
- The failed table name and specific error message are printed upon failure

#### 6. Auto-split for Large Exports

| Threshold | Behavior |
|-----------|----------|
| 50,000 rows / sheet | Auto-creates new sheets (e.g., `sheet0`, `sheet1`...) |
| 500,000 rows / workbook | Auto-creates new files (e.g., `xxx_w002.xlsx`, `xxx_w003.xlsx`...) |
| 500,000 rows / table | Enables multi-process segmented export (up to 4 processes) |

---

### Smart Pagination Strategy

| Strategy | Condition | Complexity | Performance |
|----------|-----------|------------|-------------|
| Integer PK Range | Single integer PK | O(1) | Fastest (2.7M rows < 2 min) |
| Keyset Pagination | Single string PK | O(1)/batch | Fast (10-50x faster than LIMIT/OFFSET) |
| LIMIT/OFFSET | No PK | O(n²) | Slow (add PK for better performance) |

### Data Type Mapping (excerpt)

| Source Type | SQLite | MySQL | PostgreSQL |
|-------------|--------|-------|------------|
| INTEGER / INT | INTEGER | INT | INTEGER |
| BIGINT | INTEGER | BIGINT | BIGINT |
| VARCHAR | TEXT | VARCHAR(255) | VARCHAR |
| TEXT | TEXT | TEXT | TEXT |
| JSON | TEXT | JSON | JSONB |
| BOOLEAN | INTEGER | TINYINT(1) | BOOLEAN |
| DATETIME | TEXT | DATETIME | TIMESTAMP |
| FLOAT | REAL | FLOAT | REAL |
| DOUBLE | REAL | DOUBLE | DOUBLE PRECISION |
| DECIMAL | REAL | DECIMAL(20,6) | DECIMAL(20,6) |
| BLOB | BLOB | BLOB | BYTEA |
| UUID | TEXT | VARCHAR(36) | UUID |

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success (or user cancelled) |
| 1 | Missing database driver |
| 2 | Parameter configuration error |
| 3 | Migration module load failure |
| 4 | Database connection failure |
| 5 | Migration/export failure (rolled back) |
| 130 | User interrupted (Ctrl+C) |
| 99 | Unexpected runtime error |

### Sample Log Output (Successful Migration)

```
============================================================
Checking Python database driver dependencies...
[INFO] Database driver sqlite3 installed ✓
[INFO] Database driver sqlite3 installed ✓
============================================================

============================================================
【Migration Plan Summary】
------------------------------------------------------------
Source: SQLITE
  Path: ./data/source.db
Target: SQLITE
  Path: ./data/backup.db
Tables: employees, orders (2 total)
Max threads: 4
Batch size: 500
============================================================

[INFO] 2026-07-18 15:00:00 Testing source database connection...
[SUCCESS] 2026-07-18 15:00:00 Source database connection OK
[INFO] 2026-07-18 15:00:00 Testing target database connection...
[SUCCESS] 2026-07-18 15:00:00 Target database connection OK
[INFO] 2026-07-18 15:00:00 Source: sqlite -> Target: sqlite
[INFO] 2026-07-18 15:00:00 Migration plan: 2 tables
[INFO] 2026-07-18 15:00:00 Starting migration for table employees, 1500 rows
[INFO] 2026-07-18 15:00:00 Table employees created/recreated in target
[INFO] 2026-07-18 15:00:00 Row count 1500 exceeds 3000, using batch processing (500/batch, 3 batches)
[INFO] 2026-07-18 15:00:01 Table employees progress: 3/3 batches
[SUCCESS] 2026-07-18 15:00:01 Table employees migrated, 0.05s elapsed, 1500 rows
[SUCCESS] 2026-07-18 15:00:01 Migration complete! 2 tables, 1700 rows, 0.06s total

============================================================
【Migration Result Summary】
------------------------------------------------------------
  ✓ employees: 1500 rows, 3 batches, 1 thread, 0.05s
  ✓ orders: 200 rows, 1 batch, 1 thread, 0.00s
------------------------------------------------------------
Total: 2 tables, 1700 rows
Total time: 0.06s
============================================================
```

### Quick Test

```bash
# 1. View help first
python run.py -h

# 2. SQLite → SQLite test (no database server needed)
python -c "import sqlite3;c=sqlite3.connect('test.db');[c.execute('CREATE TABLE t'+str(i)+'(id INTEGER PRIMARY KEY, name TEXT)') for i in range(3)];[c.execute('INSERT INTO t'+str(i)+'(name) VALUES(\"name_\")') for i in range(3) for _ in range(100)];c.commit();c.close();print('test.db created')"
python run.py migrate --source-type sqlite --source-path test.db --target-type sqlite --target-path test_backup.db -y

# 3. Export test
python run.py export --source-type sqlite --source-path test.db --format csv -y
```

Verify results:
```bash
python -c "import sqlite3;s=sqlite3.connect('test.db');d=sqlite3.connect('test_backup.db');[print(f't{i}: src={s.execute(\"SELECT COUNT(*) FROM t\"+str(i)).fetchone()[0]}, dst={d.execute(\"SELECT COUNT(*) FROM t\"+str(i)).fetchone()[0]}') for i in range(3)]"
```

### FAQ

**Q: "Missing database driver" error?**
A: Install the driver: `pip install pymysql` (MySQL), `pip install psycopg2-binary` (PostgreSQL), `pip install xlsxwriter` (Excel).

**Q: "Access denied" for MySQL?**
A: Check username, password, port, and ensure the user has CREATE TABLE and INSERT privileges on the target database.

**Q: Will large tables cause high memory usage?**
A: No. Tables with 3,000+ rows are automatically batched, keeping memory usage consistently low.

**Q: Why can't I see my PostgreSQL tables?**
A: The tool reads tables from the `public` schema by default. Specify the schema in `--source-database` if needed.

**Q: Why is the data order different in the target table?**
A: Migration does not guarantee row order (SQL itself does not guarantee order), but the data content and row count are identical.

**Q: Excel export fails with "need xlsxwriter or openpyxl"?**
A: Install either: `pip install xlsxwriter` (recommended, faster) or `pip install openpyxl`.

**Q: How to migrate a source table to a differently-named target table?**
A: Use `--rename`: `--rename users:t_users,orders:t_orders`.

**Q: Will exporting a table with tens of millions of rows cause an OOM?**
A: No. The tool uses streaming reads + batch writes with auto-split at 50K rows/sheet and 500K rows/workbook, keeping memory usage under control at all times.
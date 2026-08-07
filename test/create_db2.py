import sqlite3
import random
import string
from tqdm import tqdm

# 数据库文件名
DB_NAME = "test_database.db"

# 表配置：表名 -> 需要生成的行数
TABLE_CONFIG = {
    "test1": 900000,
    "test2": 2700000
}

# 字段数量（除了id之外）
FIELD_COUNT = 20

# 每次批量插入的行数
BATCH_SIZE = 5000


def generate_random_string(length=50):
    """生成指定长度的随机字符串"""
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))


def create_table(cursor, table_name):
    """创建包含 id 和 20 个随机字段的表"""
    columns = ", ".join([f"col_{i} TEXT" for i in range(1, FIELD_COUNT + 1)])
    sql = f"CREATE TABLE IF NOT EXISTS {table_name} (id INTEGER PRIMARY KEY, {columns})"
    cursor.execute(sql)


def insert_data(cursor, table_name, total_rows):
    """分批插入数据，带进度条"""
    columns = ", ".join([f"col_{i}" for i in range(1, FIELD_COUNT + 1)])
    placeholders = ", ".join(["?"] * FIELD_COUNT)

    sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

    current_id = 1
    with tqdm(total=total_rows, desc=f"生成 {table_name}", unit="行", ncols=80) as pbar:
        while current_id <= total_rows:
            # 准备一批数据 (每次最多 5000 行)
            data_batch = []
            for _ in range(min(BATCH_SIZE, total_rows - current_id + 1)):
                row = [generate_random_string(50) for _ in range(FIELD_COUNT)]
                data_batch.append(row)

            # 批量写入
            cursor.executemany(sql, data_batch)
            current_id += len(data_batch)
            pbar.update(len(data_batch))

        # 提交当前表的事务
        cursor.connection.commit()


def main():
    print(f"正在初始化数据库: {DB_NAME} ...")
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # 优化 SQLite 写入性能（生成测试数据时必开）
    cursor.execute("PRAGMA journal_mode = OFF")
    cursor.execute("PRAGMA synchronous = OFF")

    for table_name, row_count in TABLE_CONFIG.items():
        print(f"\n开始处理表: {table_name} (目标: {row_count:,} 行)")
        create_table(cursor, table_name)
        insert_data(cursor, table_name, row_count)
        print(f"✅ {table_name} 生成完毕！")

    cursor.close()
    conn.close()
    print("\n🎉 所有数据生成完成！")


if __name__ == "__main__":
    main()
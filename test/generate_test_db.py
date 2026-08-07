# -*- coding: utf-8 -*-
"""
生成测试用的 SQLite 数据库
两张表:
    test1: 70 万行 (id, uuid)
    test2: 250 万行 (id, uuid)
所有 uuid 均为唯一值
"""
import os
import sqlite3
import sys
import time
import uuid


def create_database(db_path: str,
                     test1_rows: int = 700000,
                     test2_rows: int = 2500000,
                     batch_size: int = 10000):
    """创建并写入测试数据库"""
    if os.path.exists(db_path):
        print(f"[WARN] 文件已存在，将覆盖: {db_path}")
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # 建表
    cur.execute("""
        CREATE TABLE test1 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid TEXT NOT NULL UNIQUE
        )
    """)
    cur.execute("""
        CREATE TABLE test2 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid TEXT NOT NULL UNIQUE
        )
    """)
    conn.commit()

    # ---------- 写入 test1 ----------
    print(f"\n[1/2] 写入 test1 ({test1_rows:,} 行) ...")
    start = time.time()
    for i in range(0, test1_rows, batch_size):
        end = min(i + batch_size, test1_rows)
        values = [(str(uuid.uuid4()),) for _ in range(end - i)]
        cur.executemany("INSERT INTO test1 (uuid) VALUES (?)", values)
        conn.commit()
        pct = end / test1_rows * 100
        elapsed = time.time() - start
        eta = (elapsed / (end / test1_rows)) - elapsed if end > 0 else 0
        sys.stdout.write(
            f"\r  test1 进度: {end:,}/{test1_rows:,} ({pct:5.1f}%)  "
            f"用时 {elapsed:5.1f}s  剩余约 {eta:5.1f}s"
        )
        sys.stdout.flush()
    t1_elapsed = time.time() - start
    print(f"\n  test1 完成！用时 {t1_elapsed:.2f}s  "
          f"({int(test1_rows / t1_elapsed)} 行/秒)")

    # ---------- 写入 test2 ----------
    print(f"\n[2/2] 写入 test2 ({test2_rows:,} 行) ...")
    start = time.time()
    for i in range(0, test2_rows, batch_size):
        end = min(i + batch_size, test2_rows)
        values = [(str(uuid.uuid4()),) for _ in range(end - i)]
        cur.executemany("INSERT INTO test2 (uuid) VALUES (?)", values)
        conn.commit()
        pct = end / test2_rows * 100
        elapsed = time.time() - start
        eta = (elapsed / (end / test2_rows)) - elapsed if end > 0 else 0
        sys.stdout.write(
            f"\r  test2 进度: {end:,}/{test2_rows:,} ({pct:5.1f}%)  "
            f"用时 {elapsed:5.1f}s  剩余约 {eta:5.1f}s"
        )
        sys.stdout.flush()
    t2_elapsed = time.time() - start
    print(f"\n  test2 完成！用时 {t2_elapsed:.2f}s  "
          f"({int(test2_rows / t2_elapsed)} 行/秒)")

    # ---------- 验证 ----------
    print("\n[验证]")
    cur.execute("SELECT COUNT(*) FROM test1")
    c1 = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM test2")
    c2 = cur.fetchone()[0]
    print(f"  test1 实际行数: {c1:,}")
    print(f"  test2 实际行数: {c2:,}")

    # 验证唯一性
    cur.execute("SELECT COUNT(DISTINCT uuid) FROM test1")
    u1 = cur.fetchone()[0]
    cur.execute("SELECT COUNT(DISTINCT uuid) FROM test2")
    u2 = cur.fetchone()[0]
    print(f"  test1 唯一 uuid: {u1:,} {'✓' if u1 == c1 else '✗ (有重复!)'}")
    print(f"  test2 唯一 uuid: {u2:,} {'✓' if u2 == c2 else '✗ (有重复!)'}")

    cur.close()
    conn.close()

    file_size = os.path.getsize(db_path) / 1024 / 1024
    print(f"\n  数据库文件: {db_path}")
    print(f"  文件大小: {file_size:.2f} MB")
    total = t1_elapsed + t2_elapsed
    print(f"  总用时: {total:.2f}s")
    print(f"  合计: {(c1 + c2):,} 行")


def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(output_dir, "../db/test_large.db")

    print("=" * 60)
    print("  大型 SQLite 测试数据库生成器")
    print("=" * 60)
    print(f"  目标文件: {db_path}")
    print(f"  test1: 700,000 行")
    print(f"  test2: 2,500,000 行")
    print(f"  字段: id (INTEGER PK), uuid (TEXT UNIQUE)")
    print("=" * 60)

    confirm = input("\n是否继续? [y/N]: ").strip().lower()
    if confirm not in ("y", "yes"):
        print("已取消")
        return

    try:
        create_database(db_path)
        print("\n✓ 全部完成！")
    except KeyboardInterrupt:
        print("\n\n[WARN] 用户中断")
        sys.exit(1)
    except Exception as exc:
        print(f"\n[ERROR] {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
"""
SQL 方言转换模块
负责将通用 SQL 语法转换为 SQLite / MySQL / PostgreSQL 各自的方言
"""

from typing import Dict, List, Tuple


class SQLDialect:
    """SQL 方言转换器"""

    _TYPE_MAPPING: Dict[str, Dict[str, str]] = {
        "sqlite": {
            "int": "INTEGER",
            "integer": "INTEGER",
            "bigint": "INTEGER",
            "smallint": "INTEGER",
            "tinyint": "INTEGER",
            "varchar": "TEXT",
            "char": "TEXT",
            "text": "TEXT",
            "longtext": "TEXT",
            "mediumtext": "TEXT",
            "json": "TEXT",
            "float": "REAL",
            "double": "REAL",
            "decimal": "REAL",
            "numeric": "REAL",
            "boolean": "INTEGER",
            "bool": "INTEGER",
            "datetime": "TEXT",
            "timestamp": "TEXT",
            "date": "TEXT",
            "time": "TEXT",
            "blob": "BLOB",
            "binary": "BLOB",
            "uuid": "TEXT",
        },
        "mysql": {
            "int": "INT",
            "integer": "INT",
            "bigint": "BIGINT",
            "smallint": "SMALLINT",
            "tinyint": "TINYINT",
            "varchar": "VARCHAR(255)",
            "char": "VARCHAR(255)",
            "text": "TEXT",
            "longtext": "LONGTEXT",
            "mediumtext": "MEDIUMTEXT",
            "json": "JSON",
            "float": "FLOAT",
            "double": "DOUBLE",
            "decimal": "DECIMAL(20,6)",
            "numeric": "DECIMAL(20,6)",
            "boolean": "TINYINT(1)",
            "bool": "TINYINT(1)",
            "datetime": "DATETIME",
            "timestamp": "TIMESTAMP",
            "date": "DATE",
            "time": "TIME",
            "blob": "BLOB",
            "binary": "BLOB",
            "uuid": "VARCHAR(36)",
        },
        "postgresql": {
            "int": "INTEGER",
            "integer": "INTEGER",
            "bigint": "BIGINT",
            "smallint": "SMALLINT",
            "tinyint": "SMALLINT",
            "varchar": "VARCHAR",
            "char": "CHAR",
            "text": "TEXT",
            "longtext": "TEXT",
            "mediumtext": "TEXT",
            "json": "JSONB",
            "float": "REAL",
            "double": "DOUBLE PRECISION",
            "decimal": "DECIMAL(20,6)",
            "numeric": "DECIMAL(20,6)",
            "boolean": "BOOLEAN",
            "bool": "BOOLEAN",
            "datetime": "TIMESTAMP",
            "timestamp": "TIMESTAMP",
            "date": "DATE",
            "time": "TIME",
            "blob": "BYTEA",
            "binary": "BYTEA",
            "uuid": "UUID",
        },
    }

    @classmethod
    def normalize_type(cls, raw_type: str, target_db: str) -> str:
        """将原始类型转换为目标数据库兼容的类型"""
        target = target_db.lower()
        if target not in cls._TYPE_MAPPING:
            target = "sqlite"

        if raw_type is None:
            return cls._TYPE_MAPPING[target]["text"]

        base_type = raw_type.lower().strip()
        clean_type = base_type.split("(")[0].strip()
        clean_type = clean_type.split()[0].strip()
        clean_type = clean_type.replace("unsigned", "").strip()

        mapping = cls._TYPE_MAPPING[target]
        if clean_type in mapping:
            return mapping[clean_type]

        if "int" in clean_type:
            return mapping.get("int", "INTEGER")
        if "char" in clean_type or "text" in clean_type or "enum" in clean_type:
            return mapping.get("text", "TEXT")
        if "float" in clean_type or "double" in clean_type or "real" in clean_type:
            return mapping.get("float", "REAL")
        if "decimal" in clean_type or "numeric" in clean_type:
            return mapping.get("decimal", "REAL")
        if "bool" in clean_type:
            return mapping.get("boolean", "INTEGER")
        if "date" in clean_type or "time" in clean_type:
            return mapping.get("datetime", "TEXT")
        if "blob" in clean_type or "binary" in clean_type or "bytea" in clean_type:
            return mapping.get("blob", "BLOB")
        if "json" in clean_type:
            return mapping.get("json", "TEXT")

        return mapping.get("text", "TEXT")

    @classmethod
    def quote_identifier(cls, identifier: str, db_type: str) -> str:
        """标识符引用：表名和列名"""
        db = db_type.lower()
        if db == "mysql":
            return f"`{identifier}`"
        if db in ("postgresql", "psql", "pg"):
            return f'"{identifier}"'
        return f'"{identifier}"'

    @classmethod
    def value_placeholder(cls, db_type: str, index: int = 0) -> str:
        """获取参数占位符"""
        db = db_type.lower()
        if db in ("postgresql", "psql", "pg"):
            return f"${index + 1}"
        return "?" if db == "sqlite" else "%s"

    @classmethod
    def build_create_table_sql(
        cls,
        table_name: str,
        schema: List[Tuple[str, str, bool, bool]],
        target_db: str,
    ) -> str:
        """
        生成建表 SQL
        schema: [(列名, 类型, 是否可空, 是否主键)]
        """
        quoted_table = cls.quote_identifier(table_name, target_db)
        column_defs = []
        primary_keys = []

        # 判断整数类型主键，用于 MySQL AUTO_INCREMENT 推断
        target = target_db.lower()
        is_mysql = target in ("mysql", "mariadb")

        def _is_integer_type(col_type: str) -> bool:
            if not col_type:
                return False
            ct = col_type.lower().strip()
            return any(k in ct for k in ("int", "integer", "bigint", "smallint", "tinyint", "mediumint"))

        pk_count = sum(1 for _, _, _, is_pk in schema if is_pk)

        for col_name, col_type, is_nullable, is_primary in schema:
            mapped_type = cls.normalize_type(col_type, target_db)
            quoted_col = cls.quote_identifier(col_name, target_db)
            # 主键列天然不可空（SQL 标准语义），特别修复 SQLite INTEGER PRIMARY KEY
            # 在 PRAGMA table_info 中 notnull=0 的情况，避免迁移到 MySQL 时出现
            # "All parts of a PRIMARY KEY must be NOT NULL" 错误
            if is_primary:
                is_nullable = False
            null_clause = "NULL" if is_nullable else "NOT NULL"

            extra = ""
            # MySQL：单列整数主键 → 加上 AUTO_INCREMENT，保持与源端语义一致
            if (is_mysql and is_primary and pk_count == 1
                    and _is_integer_type(col_type)):
                extra = " AUTO_INCREMENT"

            column_defs.append(f"  {quoted_col} {mapped_type} {null_clause}{extra}")
            if is_primary:
                primary_keys.append(quoted_col)

        if primary_keys:
            pk_clause = f"  PRIMARY KEY ({', '.join(primary_keys)})"
            column_defs.append(pk_clause)

        sql = f"CREATE TABLE IF NOT EXISTS {quoted_table} (\n"
        sql += ",\n".join(column_defs)
        sql += "\n);"
        return sql

    @classmethod
    def build_drop_table_sql(cls, table_name: str, db_type: str) -> str:
        """生成 DROP TABLE SQL"""
        quoted_table = cls.quote_identifier(table_name, db_type)
        return f"DROP TABLE IF EXISTS {quoted_table};"

    @classmethod
    def build_insert_sql(
        cls,
        table_name: str,
        columns: List[str],
        db_type: str,
    ) -> str:
        """生成 INSERT SQL"""
        quoted_table = cls.quote_identifier(table_name, db_type)
        quoted_columns = [cls.quote_identifier(c, db_type) for c in columns]
        placeholders = [cls.value_placeholder(db_type, i) for i in range(len(columns))]

        return (
            f"INSERT INTO {quoted_table} ({', '.join(quoted_columns)}) "
            f"VALUES ({', '.join(placeholders)});"
        )

    @classmethod
    def build_select_batched_sql(
        cls,
        table_name: str,
        columns: List[str],
        batch_size: int,
        offset: int,
        db_type: str,
    ) -> str:
        """生成分批读取数据的 SELECT SQL"""
        quoted_table = cls.quote_identifier(table_name, db_type)
        quoted_columns = [cls.quote_identifier(c, db_type) for c in columns]
        col_sql = ", ".join(quoted_columns)

        if db_type.lower() in ("postgresql", "psql", "pg"):
            return (
                f"SELECT {col_sql} FROM {quoted_table} "
                f"LIMIT {batch_size} OFFSET {offset};"
            )
        if db_type.lower() == "mysql":
            return (
                f"SELECT {col_sql} FROM {quoted_table} "
                f"LIMIT {offset}, {batch_size};"
            )
        return (
            f"SELECT {col_sql} FROM {quoted_table} "
            f"LIMIT {batch_size} OFFSET {offset};"
        )
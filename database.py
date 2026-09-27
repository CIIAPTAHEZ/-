"""
Модуль работы с базой данных SQLite.

Содержит схему БД, функции инициализации, резервного копирования
и журналирования изменений.
"""
import os
import shutil
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Iterator

import config

SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS claims (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    claim_number INTEGER UNIQUE NOT NULL,
    date_received TEXT,
    contact_method TEXT,
    applicant TEXT,
    address TEXT,
    contacts TEXT,
    product_name TEXT,
    defect TEXT,
    quantity INTEGER,
    status TEXT,
    decision_date TEXT,
    shipment_date TEXT,
    return_date TEXT,
    folder_path TEXT,
    is_production_defect INTEGER DEFAULT 0,
    comment TEXT
);

CREATE TABLE IF NOT EXISTS operational_control (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    control_date TEXT,
    batch TEXT,
    press_operator TEXT,
    mold TEXT,
    shift_master TEXT,
    defect TEXT,
    quantity INTEGER,
    check_time TEXT
);

CREATE TABLE IF NOT EXISTS micro_lift_forms (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS micro_lift (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    form_id INTEGER NOT NULL,
    measure_type TEXT NOT NULL DEFAULT 'seat',
    time_seconds INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    comment TEXT
);

CREATE TABLE IF NOT EXISTS change_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    log_date TEXT NOT NULL,
    table_name TEXT NOT NULL,
    record_id INTEGER,
    action TEXT NOT NULL,
    details TEXT
);
"""


@contextmanager
def get_connection() -> Iterator[sqlite3.Connection]:
    """Контекстный менеджер подключения к БД.

    Гарантирует выполнение транзакции и закрытие соединения.
    """
    conn = sqlite3.connect(config.DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    """Инициализирует базу данных: создаёт таблицы при первом запуске."""
    config.ensure_directories()
    with get_connection() as conn:
        conn.executescript(SCHEMA)
        # Миграция: добавляем колонки в существующие БД
        cols = [row["name"] for row in conn.execute("PRAGMA table_info(claims)").fetchall()]
        if "is_production_defect" not in cols:
            conn.execute(
                "ALTER TABLE claims ADD COLUMN is_production_defect INTEGER DEFAULT 0")
        if "comment" not in cols:
            conn.execute(
                "ALTER TABLE claims ADD COLUMN comment TEXT")

        # Миграция: убираем UNIQUE с колонки batch в operational_control.
        # SQLite не умеет удалять ограничения через ALTER TABLE,
        # поэтому пересоздаём таблицу без UNIQUE.
        sql = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' "
            "AND name='operational_control'").fetchone()
        has_unique_legacy = sql and "UNIQUE" in (sql["sql"] or "").upper()

        if has_unique_legacy:
            conn.execute("""
                CREATE TABLE operational_control_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    control_date TEXT,
                    batch TEXT,
                    press_operator TEXT,
                    mold TEXT,
                    shift_master TEXT,
                    defect TEXT,
                    quantity INTEGER,
                    check_time TEXT
                )
            """)
            conn.execute("""
                INSERT INTO operational_control_new
                    (id, control_date, batch, press_operator, mold,
                     shift_master, defect, quantity, check_time)
                SELECT id, control_date, batch, press_operator, mold,
                       shift_master, defect, quantity, check_time
                FROM operational_control
            """)
            conn.execute("DROP TABLE operational_control")
            conn.execute("ALTER TABLE operational_control_new RENAME TO operational_control")

        # Миграция микролифта:
        # 1) если нет колонки measure_type — добавляем (старым записям ставим 'seat')
        cols = [row["name"] for row in conn.execute("PRAGMA table_info(micro_lift)").fetchall()]
        if "measure_type" not in cols:
            conn.execute(
                "ALTER TABLE micro_lift ADD COLUMN measure_type TEXT NOT NULL DEFAULT 'seat'")

        # 2) если таблица форм пустая — заполняем дефолтным списком
        #    (это же сохраняет привязку старых записей form_id=1..21 к «Форма 01»…«Форма 21»)
        count_forms = conn.execute(
            "SELECT COUNT(*) FROM micro_lift_forms").fetchone()[0]
        if count_forms == 0:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for i, name in enumerate(config.MICROLIFT_FORMS, start=1):
                conn.execute(
                    "INSERT OR IGNORE INTO micro_lift_forms (id, name, created_at) "
                    "VALUES (?, ?, ?)",
                    (i, name, now))


def create_backup() -> str:
    """Создаёт резервную копию файла БД в папке backups.

    Возвращает путь к созданной копии.
    """
    config.ensure_directories()
    if not config.DB_PATH.exists():
        return ""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = config.BACKUP_DIR / f"claims_backup_{timestamp}.db"

    # Используем API SQLite для создания согласованной копии
    src = sqlite3.connect(config.DB_PATH)
    dst = sqlite3.connect(backup_path)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    return str(backup_path)


def log_change(table_name: str, record_id: int | None,
               action: str, details: str = "",
               conn: sqlite3.Connection | None = None) -> None:
    """Записывает действие пользователя в журнал изменений.

    Если передано соединение conn, запись выполняется в его транзакции.
    Иначе открывается отдельное соединение.
    """
    params = (datetime.now().isoformat(sep=" "), table_name, record_id, action, details)
    if conn is not None:
        conn.execute(
            "INSERT INTO change_log (log_date, table_name, record_id, action, details) "
            "VALUES (?, ?, ?, ?, ?)",
            params,
        )
    else:
        with get_connection() as c:
            c.execute(
                "INSERT INTO change_log (log_date, table_name, record_id, action, details) "
                "VALUES (?, ?, ?, ?, ?)",
                params,
            )

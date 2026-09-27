"""
Flask-приложение «Анализ рекламаций и операционного контроля».

Предоставляет REST API и раздаёт статический интерфейс.
"""
import os
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from flask import (Flask, jsonify, request, send_file, send_from_directory)
from werkzeug.utils import secure_filename

import config
import importer
import reports
from database import create_backup, get_connection, init_db, log_change

# Определяем папку со статическими файлами.
# В упакованном приложении (PyInstaller) static извлекается во временную папку _MEIPASS.
_STATIC_FOLDER = "static"
if getattr(sys, "frozen", False):
    _STATIC_FOLDER = os.path.join(sys._MEIPASS, "static")

app = Flask(__name__, static_folder=_STATIC_FOLDER, static_url_path="/static")

# Разрешённые типы файлов для прикрепления
ALLOWED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".doc", ".docx", ".jpg",
                      ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".zip", ".rar"}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 МБ


# ---------------------------------------------------------------------------
# Служебные функции
# ---------------------------------------------------------------------------

def _query_to_json(query: str, params: tuple = ()) -> list[dict[str, Any]]:
    """Выполняет SELECT и возвращает список словарей."""
    with get_connection() as conn:
        rows = conn.execute(query, params).fetchall()
    return [dict(row) for row in rows]


def _apply_filters(table: str, filters: dict[str, Any]) -> tuple[str, list]:
    """Формирует SQL-условия WHERE по переданным фильтрам.

    Поддерживает: search, date_from, date_to, точное совпадение по полям.
    """
    where: list[str] = []
    params: list = []

    search = filters.get("search", "").strip()
    if search:
        searchable = {
            "claims": ["claim_number", "contact_method", "applicant", "address",
                       "contacts", "product_name", "defect", "status"],
            "control": ["batch", "press_operator", "mold", "shift_master", "defect"],
        }.get(table, [])
        like_clauses = " OR ".join(f"CAST({field} AS TEXT) LIKE ?" for field in searchable)
        like_params = [f"%{search}%"] * len(searchable)
        where.append(f"({like_clauses})")
        params.extend(like_params)

    date_from = filters.get("date_from", "").strip()
    date_to = filters.get("date_to", "").strip()
    date_field = "date_received" if table == "claims" else "control_date"
    if date_from:
        where.append(f"{date_field} >= ?")
        params.append(date_from)
    if date_to:
        where.append(f"{date_field} <= ?")
        params.append(date_to)

    # Точные фильтры по полям (передаются как filters[f"eq_{field}"])
    for key, value in filters.items():
        if key.startswith("eq_") and value:
            field = key[3:]
            if field in _table_fields(table):
                where.append(f"{field} = ?")
                params.append(value)

    # Фильтры "входит в список" (передаются как filters[f"in_{field}"] = "a,b,c")
    for key, value in filters.items():
        if key.startswith("in_") and value:
            field = key[3:]
            if field in _table_fields(table):
                values = [v.strip() for v in str(value).split(",") if v.strip()]
                if values:
                    placeholders = ", ".join("?" for _ in values)
                    where.append(f"{field} IN ({placeholders})")
                    params.extend(values)

    # Числовые диапазоны (передаются как filters[f"min_{field}"] / filters[f"max_{field}"])
    for key, value in filters.items():
        if key.startswith("min_") and value:
            field = key[4:]
            if field in _table_fields(table):
                where.append(f"CAST({field} AS INTEGER) >= ?")
                params.append(value)
    for key, value in filters.items():
        if key.startswith("max_") and value:
            field = key[4:]
            if field in _table_fields(table):
                where.append(f"CAST({field} AS INTEGER) <= ?")
                params.append(value)

    where_sql = f" WHERE {' AND '.join(where)}" if where else ""
    return where_sql, params


def _table_fields(table: str) -> set[str]:
    """Возвращает допустимые поля для фильтрации по таблице."""
    if table == "claims":
        return {"claim_number", "date_received", "contact_method", "applicant",
                "address", "contacts", "product_name", "defect", "quantity",
                "status", "decision_date", "shipment_date", "return_date",
                "is_production_defect", "comment"}
    return {"control_date", "batch", "press_operator", "mold", "shift_master",
            "defect", "quantity", "check_time"}


def _table_name(table: str) -> str:
    """Возвращает имя таблицы БД по типу."""
    if table == "claims":
        return "claims"
    return "operational_control"


def _validate_claim_data(data: dict[str, Any]) -> str | None:
    """Проверяет данные претензии. Возвращает сообщение об ошибке или None."""
    if "claim_number" not in data or data.get("claim_number") in (None, ""):
        return "Номер претензии обязателен для заполнения."
    try:
        int(data["claim_number"])
    except (TypeError, ValueError):
        return "Номер претензии должен быть целым числом."
    return None


def _validate_control_data(data: dict[str, Any]) -> str | None:
    """Проверяет данные операционного контроля."""
    if "batch" not in data or not str(data.get("batch") or "").strip():
        return "Номер партии обязателен для заполнения."
    return None


# ---------------------------------------------------------------------------
# Главная страница
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


# ---------------------------------------------------------------------------
# Импорт из Excel
# ---------------------------------------------------------------------------

@app.route("/api/import", methods=["POST"])
def import_excel():
    """Импортирует данные из загруженного XLSX-файла."""
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "Файл не выбран."}), 400

    # Определяем расширение по оригинальному имени файла (secure_filename
    # удаляет кириллицу, из-за чего расширение может потеряться)
    original_name = file.filename or ""
    ext = Path(original_name).suffix.lower()
    if ext not in (".xlsx", ".xls"):
        return jsonify({"error": "Поддерживаются только файлы формата .xlsx и .xls"}), 400

    temp_path = config.EXPORT_DIR / f"upload_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{ext}"
    file.save(temp_path)

    try:
        # Пытаемся импортировать оба листа, если они присутствуют
        result: dict[str, Any] = {}
        try:
            result["claims"] = importer.import_claims(temp_path)
        except importer.ImportError as exc:
            result["claims_error"] = str(exc)
        try:
            result["control"] = importer.import_operational_control(temp_path)
        except importer.ImportError as exc:
            result["control_error"] = str(exc)

        if "claims_error" in result and "control_error" in result:
            return jsonify({
                "error": f"Не удалось импортировать данные: {result['claims_error']}; "
                         f"{result['control_error']}"
            }), 400

        return jsonify({
            "message": "Импорт завершён.",
            "claims": result.get("claims", {"added": 0, "updated": 0, "skipped": 0}),
            "control": result.get("control", {"added": 0, "updated": 0, "skipped": 0}),
        })
    finally:
        if temp_path.exists():
            temp_path.unlink()


# ---------------------------------------------------------------------------
# Претензии
# ---------------------------------------------------------------------------

@app.route("/api/claims", methods=["GET"])
def get_claims():
    """Возвращает список претензий с фильтрацией, сортировкой и пагинацией."""
    args = request.args
    filters = {
        "search": args.get("search", ""),
        "date_from": args.get("date_from", ""),
        "date_to": args.get("date_to", ""),
    }
    for key, value in args.items():
        if key.startswith("eq_") or key.startswith("min_") or key.startswith("max_") or key.startswith("in_"):
            filters[key] = value

    where_sql, params = _apply_filters("claims", filters)

    # Многоуровневая сортировка: sort_by=поле;поле, sort_dir=asc;desc
    sort_fields = [s.strip() for s in args.get("sort_by", "date_received").split(";") if s.strip()]
    sort_dirs = [s.strip() for s in args.get("sort_dir", "desc").split(";") if s.strip()]
    order_parts = []
    for i, field in enumerate(sort_fields):
        if field not in _table_fields("claims"):
            continue
        direction = "ASC" if i < len(sort_dirs) and sort_dirs[i] == "asc" else "DESC"
        if not direction:
            direction = "DESC"
        order_parts.append(f"{field} {direction}")
    if not order_parts:
        order_parts = ["date_received DESC"]
    order_sql = ", ".join(order_parts)

    # Общее количество
    count_row = _query_to_json(
        f"SELECT COUNT(*) as cnt FROM claims{where_sql}", tuple(params)
    )[0]
    total = count_row["cnt"]

    # Пагинация
    page = max(int(args.get("page", 1)), 1)
    page_size = min(max(int(args.get("page_size", 50)), 1), 200)
    offset = (page - 1) * page_size

    rows = _query_to_json(
        f"""SELECT * FROM claims{where_sql}
            ORDER BY {order_sql}
            LIMIT ? OFFSET ?""",
        tuple(params) + (page_size, offset),
    )

    # Сбор уникальных значений для фильтров по всей таблице
    # (независимо от пагинации и текущих фильтров), чтобы список
    # в фильтре всегда был полным — как в Excel.
    column_values: dict[str, list[str]] = {}
    for col in ["contact_method", "applicant", "address", "product_name", "defect", "status", "shipment_date"]:
        extra = f"{col} IS NOT NULL AND TRIM({col}) != ''"
        cond_sql = f" WHERE {extra}"
        col_rows = _query_to_json(
            f"SELECT DISTINCT {col} as v FROM claims{cond_sql} ORDER BY v",
        )
        column_values[col] = [r["v"] for r in col_rows if r["v"]]

    return jsonify({
        "data": rows, "total": total, "page": page, "page_size": page_size,
        "column_values": column_values,
    })


@app.route("/api/claims/<int:claim_id>", methods=["GET"])
def get_claim(claim_id: int):
    rows = _query_to_json("SELECT * FROM claims WHERE id = ?", (claim_id,))
    if not rows:
        return jsonify({"error": "Запись не найдена."}), 404
    return jsonify(rows[0])


@app.route("/api/claims", methods=["POST"])
def create_claim():
    """Создаёт новую претензию и папку для файлов."""
    data = request.get_json(silent=True) or {}
    error = _validate_claim_data(data)
    if error:
        return jsonify({"error": error}), 400

    claim_number = int(data["claim_number"])
    date_received = data.get("date_received") or None

    folder_path = importer._build_folder_path(claim_number, date_received)
    Path(folder_path).mkdir(parents=True, exist_ok=True)

    with get_connection() as conn:
        cur = conn.execute(
            """INSERT INTO claims
               (claim_number, date_received, contact_method, applicant,
                address, contacts, product_name, defect, quantity,
                status, decision_date, shipment_date, return_date, folder_path,
                is_production_defect, comment)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (claim_number, date_received,
             data.get("contact_method"), data.get("applicant"),
             data.get("address"), data.get("contacts"),
             data.get("product_name"), data.get("defect"),
             data.get("quantity"), data.get("status"),
             data.get("decision_date"), data.get("shipment_date"),
             data.get("return_date"), folder_path,
             1 if data.get("is_production_defect") else 0,
             data.get("comment")),
        )
        record_id = cur.lastrowid
        log_change("claims", record_id, "insert",
                   f"Создание: претензия {claim_number}", conn=conn)

    return jsonify({"message": "Претензия добавлена.", "id": record_id}), 201


@app.route("/api/claims/<int:claim_id>", methods=["PUT"])
def update_claim(claim_id: int):
    """Обновляет претензию."""
    data = request.get_json(silent=True) or {}
    error = _validate_claim_data(data)
    if error:
        return jsonify({"error": error}), 400

    with get_connection() as conn:
        existing = conn.execute("SELECT id FROM claims WHERE id = ?", (claim_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Запись не найдена."}), 404

        conn.execute(
            """UPDATE claims SET
                claim_number = ?, date_received = ?, contact_method = ?, applicant = ?,
                address = ?, contacts = ?, product_name = ?, defect = ?,
                quantity = ?, status = ?, decision_date = ?,
                shipment_date = ?, return_date = ?, is_production_defect = ?,
                comment = ?
               WHERE id = ?""",
            (int(data["claim_number"]), data.get("date_received"),
             data.get("contact_method"), data.get("applicant"),
             data.get("address"), data.get("contacts"),
             data.get("product_name"), data.get("defect"),
             data.get("quantity"), data.get("status"),
             data.get("decision_date"), data.get("shipment_date"),
             data.get("return_date"),
             1 if data.get("is_production_defect") else 0,
             data.get("comment"), claim_id),
        )
        log_change("claims", claim_id, "update",
                   f"Редактирование: претензия {data['claim_number']}", conn=conn)

    return jsonify({"message": "Претензия обновлена."})


@app.route("/api/claims/<int:claim_id>/defect", methods=["POST"])
def toggle_production_defect(claim_id: int):
    """Переключает признак «Производственный брак» у претензии."""
    data = request.get_json(silent=True) or {}
    value = 1 if data.get("is_production_defect") else 0
    with get_connection() as conn:
        existing = conn.execute(
            "SELECT claim_number FROM claims WHERE id = ?", (claim_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Запись не найдена."}), 404
        conn.execute(
            "UPDATE claims SET is_production_defect = ? WHERE id = ?",
            (value, claim_id))
        log_change("claims", claim_id, "update",
                   f"Производственный брак: {'да' if value else 'нет'} "
                   f"(претензия {existing['claim_number']})", conn=conn)
    return jsonify({"message": "Признак обновлён.", "is_production_defect": value})


@app.route("/api/claims/<int:claim_id>", methods=["DELETE"])
def delete_claim(claim_id: int):
    with get_connection() as conn:
        existing = conn.execute("SELECT claim_number FROM claims WHERE id = ?", (claim_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Запись не найдена."}), 404
        conn.execute("DELETE FROM claims WHERE id = ?", (claim_id,))
        log_change("claims", claim_id, "delete",
                   f"Удаление: претензия {existing['claim_number']}", conn=conn)
    return jsonify({"message": "Претензия удалена."})


# ---------------------------------------------------------------------------
# Операционный контроль
# ---------------------------------------------------------------------------

@app.route("/api/control", methods=["GET"])
def get_control():
    """Возвращает список записей операционного контроля."""
    args = request.args
    filters = {
        "search": args.get("search", ""),
        "date_from": args.get("date_from", ""),
        "date_to": args.get("date_to", ""),
    }
    for key, value in args.items():
        if key.startswith("eq_") or key.startswith("min_") or key.startswith("max_") or key.startswith("in_"):
            filters[key] = value

    where_sql, params = _apply_filters("control", filters)

    # Многоуровневая сортировка
    sort_fields = [s.strip() for s in args.get("sort_by", "control_date").split(";") if s.strip()]
    sort_dirs = [s.strip() for s in args.get("sort_dir", "desc").split(";") if s.strip()]
    order_parts = []
    for i, field in enumerate(sort_fields):
        if field not in _table_fields("control"):
            continue
        direction = "ASC" if i < len(sort_dirs) and sort_dirs[i] == "asc" else "DESC"
        order_parts.append(f"{field} {direction}")
    if not order_parts:
        order_parts = ["control_date DESC"]
    order_sql = ", ".join(order_parts)

    count_row = _query_to_json(
        f"SELECT COUNT(*) as cnt FROM operational_control{where_sql}", tuple(params)
    )[0]
    total = count_row["cnt"]

    page = max(int(args.get("page", 1)), 1)
    page_size = min(max(int(args.get("page_size", 50)), 1), 200)
    offset = (page - 1) * page_size

    rows = _query_to_json(
        f"""SELECT * FROM operational_control{where_sql}
            ORDER BY {order_sql}
            LIMIT ? OFFSET ?""",
        tuple(params) + (page_size, offset),
    )

    # Сбор уникальных значений для фильтров по всей таблице
    # (независимо от пагинации и текущих фильтров), чтобы список
    # в фильтре всегда был полным — как в Excel.
    column_values: dict[str, list[str]] = {}
    for col in ["batch", "press_operator", "mold", "shift_master", "defect", "check_time"]:
        extra = f"{col} IS NOT NULL AND TRIM({col}) != ''"
        cond_sql = f" WHERE {extra}"
        col_rows = _query_to_json(
            f"SELECT DISTINCT {col} as v FROM operational_control{cond_sql} ORDER BY v",
        )
        column_values[col] = [r["v"] for r in col_rows if r["v"]]

    return jsonify({
        "data": rows, "total": total, "page": page, "page_size": page_size,
        "column_values": column_values,
    })


@app.route("/api/control/<int:record_id>", methods=["GET"])
def get_control_record(record_id: int):
    rows = _query_to_json(
        "SELECT * FROM operational_control WHERE id = ?", (record_id,))
    if not rows:
        return jsonify({"error": "Запись не найдена."}), 404
    return jsonify(rows[0])


@app.route("/api/control", methods=["POST"])
def create_control():
    data = request.get_json(silent=True) or {}
    error = _validate_control_data(data)
    if error:
        return jsonify({"error": error}), 400

    with get_connection() as conn:
        cur = conn.execute(
            """INSERT INTO operational_control
               (control_date, batch, press_operator, mold, shift_master,
                defect, quantity, check_time)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (data.get("control_date"), data.get("batch"),
             data.get("press_operator"), data.get("mold"),
             data.get("shift_master"), data.get("defect"),
             data.get("quantity"), data.get("check_time")),
        )
        record_id = cur.lastrowid
        log_change("operational_control", record_id, "insert",
                   f"Создание: партия {data['batch']}", conn=conn)

    return jsonify({"message": "Запись добавлена.", "id": record_id}), 201


@app.route("/api/control/<int:record_id>", methods=["PUT"])
def update_control(record_id: int):
    data = request.get_json(silent=True) or {}
    error = _validate_control_data(data)
    if error:
        return jsonify({"error": error}), 400

    with get_connection() as conn:
        existing = conn.execute(
            "SELECT id FROM operational_control WHERE id = ?", (record_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Запись не найдена."}), 404

        conn.execute(
            """UPDATE operational_control SET
                control_date = ?, batch = ?, press_operator = ?, mold = ?,
                shift_master = ?, defect = ?, quantity = ?, check_time = ?
               WHERE id = ?""",
            (data.get("control_date"), data.get("batch"),
             data.get("press_operator"), data.get("mold"),
             data.get("shift_master"), data.get("defect"),
             data.get("quantity"), data.get("check_time"), record_id),
        )
        log_change("operational_control", record_id, "update",
                   f"Редактирование: партия {data['batch']}", conn=conn)

    return jsonify({"message": "Запись обновлена."})


@app.route("/api/control/<int:record_id>", methods=["DELETE"])
def delete_control(record_id: int):
    with get_connection() as conn:
        existing = conn.execute(
            "SELECT batch FROM operational_control WHERE id = ?", (record_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Запись не найдена."}), 404
        conn.execute("DELETE FROM operational_control WHERE id = ?", (record_id,))
        log_change("operational_control", record_id, "delete",
                   f"Удаление: партия {existing['batch']}", conn=conn)
    return jsonify({"message": "Запись удалена."})


# ---------------------------------------------------------------------------
# Микролифт (время опускания)
# ---------------------------------------------------------------------------

MEASURE_TYPES_ALLOWED = set(config.MEASURE_TYPES.keys())


def _get_microlift_form_row(form_id: Any, conn):
    """Возвращает строку формы по id или None."""
    return conn.execute(
        "SELECT id, name FROM micro_lift_forms WHERE id = ?", (form_id,)).fetchone()


@app.route("/api/microlift/forms", methods=["GET"])
def get_microlift_forms():
    """Возвращает список форм изделий из БД (пользовательские названия)."""
    rows = _query_to_json(
        "SELECT id, name, created_at FROM micro_lift_forms "
        "ORDER BY name COLLATE NOCASE, id")
    return jsonify({"forms": rows})


@app.route("/api/microlift/forms", methods=["POST"])
def create_microlift_form():
    """Создаёт новую форму изделия с произвольным названием."""
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", "") or "").strip()
    if not name:
        return jsonify({"error": "Укажите название формы."}), 400
    if len(name) > 100:
        return jsonify({"error": "Название формы слишком длинное (до 100 символов)."}), 400
    with get_connection() as conn:
        existing = conn.execute(
            "SELECT id FROM micro_lift_forms WHERE name = ? COLLATE NOCASE",
            (name,)).fetchone()
        if existing:
            return jsonify({"error": "Форма с таким названием уже существует."}), 409
        cur = conn.execute(
            "INSERT INTO micro_lift_forms (name, created_at) VALUES (?, ?)",
            (name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        form_id = cur.lastrowid
        log_change("micro_lift_forms", form_id, "insert",
                   f"Создана форма: {name}", conn=conn)
    return jsonify({"message": "Форма создана.", "id": form_id}), 201


@app.route("/api/microlift/forms/<int:form_id>", methods=["PUT"])
def rename_microlift_form(form_id: int):
    """Переименовывает форму изделия."""
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", "") or "").strip()
    if not name:
        return jsonify({"error": "Укажите название формы."}), 400
    if len(name) > 100:
        return jsonify({"error": "Название формы слишком длинное (до 100 символов)."}), 400
    with get_connection() as conn:
        row = _get_microlift_form_row(form_id, conn)
        if not row:
            return jsonify({"error": "Форма не найдена."}), 404
        existing = conn.execute(
            "SELECT id FROM micro_lift_forms "
            "WHERE name = ? COLLATE NOCASE AND id <> ?",
            (name, form_id)).fetchone()
        if existing:
            return jsonify({"error": "Форма с таким названием уже существует."}), 409
        conn.execute("UPDATE micro_lift_forms SET name = ? WHERE id = ?", (name, form_id))
        log_change("micro_lift_forms", form_id, "update",
                   f"Форма «{row['name']}» → «{name}»", conn=conn)
    return jsonify({"message": "Форма переименована."})


@app.route("/api/microlift/forms/<int:form_id>", methods=["DELETE"])
def delete_microlift_form(form_id: int):
    """Удаляет форму вместе с её замерами (необратимо)."""
    with get_connection() as conn:
        row = _get_microlift_form_row(form_id, conn)
        if not row:
            return jsonify({"error": "Форма не найдена."}), 404
        n = conn.execute(
            "SELECT COUNT(*) FROM micro_lift WHERE form_id = ?",
            (form_id,)).fetchone()[0]
        conn.execute("DELETE FROM micro_lift WHERE form_id = ?", (form_id,))
        conn.execute("DELETE FROM micro_lift_forms WHERE id = ?", (form_id,))
        log_change("micro_lift_forms", form_id, "delete",
                   f"Удалена форма «{row['name']}» и её замеры ({n})", conn=conn)
    return jsonify({"message": "Форма удалена вместе с замерами."})


@app.route("/api/microlift/stats", methods=["GET"])
def get_microlift_stats():
    """Возвращает статистику по формам и типам замеров.

    Каждый элемент = {form_id, name, measure_type, n, avg_time,
                      min_time, max_time, range}. Для пустых комбинаций
                      форма×тип n=0 и None-значения.
    """
    rows = _query_to_json(
        """SELECT form_id, measure_type,
                  COUNT(*) as n,
                  AVG(time_seconds) as avg_time,
                  MIN(time_seconds) as min_time,
                  MAX(time_seconds) as max_time
           FROM micro_lift
           GROUP BY form_id, measure_type""",
    )
    by_form_type: dict[tuple[int, str], dict[str, Any]] = {}
    total_records = 0
    for row in rows:
        key = (int(row["form_id"]), row["measure_type"])
        by_form_type[key] = row
        total_records += row["n"]

    forms_db = _query_to_json(
        "SELECT id, name FROM micro_lift_forms ORDER BY name COLLATE NOCASE, id")

    result: list[dict[str, Any]] = []
    for f in forms_db:
        form_id = int(f["id"])
        for mtype in config.MEASURE_TYPES:
            row = by_form_type.get((form_id, mtype))
            if row:
                avg = row["avg_time"]
                min_v = row["min_time"]
                max_v = row["max_time"]
                result.append({
                    "form_id": form_id,
                    "name": f["name"],
                    "measure_type": mtype,
                    "measure_type_label": config.MEASURE_TYPES[mtype],
                    "n": row["n"],
                    "avg_time": round(float(avg), 1) if avg is not None else None,
                    "min_time": min_v,
                    "max_time": max_v,
                    "range": (max_v or 0) - (min_v or 0),
                })
            else:
                result.append({
                    "form_id": form_id,
                    "name": f["name"],
                    "measure_type": mtype,
                    "measure_type_label": config.MEASURE_TYPES[mtype],
                    "n": 0,
                    "avg_time": None,
                    "min_time": None,
                    "max_time": None,
                    "range": None,
                })
    return jsonify({"forms": result, "total_records": total_records})


@app.route("/api/microlift", methods=["GET"])
def get_microlift_history():
    """Возвращает историю замеров для формы (+ опциональный тип замера)."""
    form_id = request.args.get("form_id", type=int)
    if form_id is None:
        return jsonify({"error": "Не указана форма изделия."}), 400

    measure_type = request.args.get("measure_type", "").strip() or None
    if measure_type is not None and measure_type not in MEASURE_TYPES_ALLOWED:
        return jsonify({"error": "Недопустимый тип замера."}), 400

    where = "WHERE form_id = ?"
    params: list = [form_id]
    if measure_type:
        where += " AND measure_type = ?"
        params.append(measure_type)

    rows = _query_to_json(
        f"""SELECT * FROM micro_lift
            {where}
            ORDER BY created_at DESC, id DESC""",
        tuple(params),
    )
    return jsonify({"data": rows, "form_id": form_id, "measure_type": measure_type})


@app.route("/api/microlift", methods=["POST"])
def create_microlift():
    """Добавляет замер времени опускания (сиденье/крышка/изделие)."""
    data = request.get_json(silent=True) or {}

    try:
        form_id = int(data.get("form_id"))
    except (TypeError, ValueError):
        return jsonify({"error": "Выберите форму изделия."}), 400

    measure_type = str(data.get("measure_type", "seat") or "seat")
    if measure_type not in MEASURE_TYPES_ALLOWED:
        return jsonify({"error": "Недопустимый тип замера."}), 400

    try:
        time_seconds = int(data.get("time_seconds"))
    except (TypeError, ValueError):
        return jsonify({"error": "Время должно быть целым числом секунд."}), 400
    if not 1 <= time_seconds <= 60:
        return jsonify({"error": "Время должно быть от 1 до 60 секунд."}), 400

    comment = str(data.get("comment", "") or "").strip()[:240]

    with get_connection() as conn:
        form_row = _get_microlift_form_row(form_id, conn)
        if not form_row:
            return jsonify({"error": "Форма изделия не найдена."}), 400
        cur = conn.execute(
            """INSERT INTO micro_lift (form_id, measure_type, time_seconds, created_at, comment)
               VALUES (?, ?, ?, ?, ?)""",
            (form_id, measure_type, time_seconds,
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"), comment),
        )
        record_id = cur.lastrowid
        log_change("micro_lift", record_id, "insert",
                   f"Форма «{form_row['name']}» "
                   f"[{config.MEASURE_TYPES[measure_type]}]: {time_seconds} с"
                   + (f" ({comment})" if comment else ""),
                   conn=conn)

    return jsonify({"message": "Замер сохранён.", "id": record_id}), 201


@app.route("/api/microlift/<int:record_id>", methods=["DELETE"])
def delete_microlift(record_id: int):
    """Удаляет замер времени опускания микролифта."""
    with get_connection() as conn:
        existing = conn.execute(
            "SELECT form_id, measure_type, time_seconds FROM micro_lift WHERE id = ?",
            (record_id,)).fetchone()
        if not existing:
            return jsonify({"error": "Запись не найдена."}), 404
        conn.execute("DELETE FROM micro_lift WHERE id = ?", (record_id,))
        form_row = _get_microlift_form_row(existing["form_id"], conn)
        form_label = form_row["name"] if form_row else f"#{existing['form_id']}"
        mtype_label = config.MEASURE_TYPES.get(existing["measure_type"], existing["measure_type"])
        log_change("micro_lift", record_id, "delete",
                   f"Форма «{form_label}» [{mtype_label}]: {existing['time_seconds']} с",
                   conn=conn)
    return jsonify({"message": "Запись удалена."})


# ---------------------------------------------------------------------------
# Аналитика
# ---------------------------------------------------------------------------

@app.route("/api/analytics", methods=["GET"])
def analytics():
    """Строит агрегированную статистику для диаграмм и сводных таблиц."""
    table_type = request.args.get("table", "claims")
    metric = request.args.get("metric", "defect")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    table = _table_name(table_type)
    date_field = "date_received" if table_type == "claims" else "control_date"

    where_parts: list[str] = []
    params: list = []
    if date_from:
        where_parts.append(f"{date_field} >= ?")
        params.append(date_from)
    if date_to:
        where_parts.append(f"{date_field} <= ?")
        params.append(date_to)
    where_sql = f" WHERE {' AND '.join(where_parts)}" if where_parts else ""

    # Проверка допустимости метрики
    allowed_metrics = {
        "claims": {"defect", "product_name", "applicant", "status", "contact_method"},
        "control": {"defect", "press_operator", "mold", "shift_master"},
    }
    if metric not in allowed_metrics.get(table_type, set()):
        return jsonify({"error": "Недопустимый показатель для выбранной таблицы."}), 400

    # Группировка по выбранному показателю
    grouped = _query_to_json(
        f"""SELECT COALESCE(NULLIF(TRIM({metric}), ''), '(не указано)') as label,
                   COUNT(*) as value
            FROM {table}{where_sql}
            GROUP BY label
            ORDER BY value DESC
            LIMIT 20""",
        tuple(params),
    )

    # Динамика по месяцам
    monthly = _query_to_json(
        f"""SELECT substr({date_field}, 1, 7) as label,
                   COUNT(*) as value
            FROM {table}{where_sql}
            GROUP BY label
            ORDER BY label""",
        tuple(params),
    )

    # Сводка по статусам для претензий
    statuses = []
    if table_type == "claims":
        statuses = _query_to_json(
            f"""SELECT COALESCE(NULLIF(TRIM(status), ''), '(не указано)') as label,
                       COUNT(*) as value
                FROM claims{where_sql}
                GROUP BY label
                ORDER BY value DESC""",
            tuple(params),
        )

    return jsonify({
        "grouped": grouped,
        "monthly": monthly,
        "statuses": statuses,
    })


@app.route("/api/crosstab", methods=["GET"])
def crosstab():
    """Строит сводную таблицу по двум измерениям."""
    table_type = request.args.get("table", "claims")
    row_field = request.args.get("row", "defect")
    col_field = request.args.get("col", "status")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    table = _table_name(table_type)
    date_field = "date_received" if table_type == "claims" else "control_date"

    allowed_fields = {
        "claims": {"defect", "product_name", "applicant", "status", "contact_method"},
        "control": {"defect", "press_operator", "mold", "shift_master"},
    }
    if row_field not in allowed_fields.get(table_type, set()) or \
       col_field not in allowed_fields.get(table_type, set()):
        return jsonify({"error": "Недопустимые поля для сводной таблицы."}), 400

    where_parts: list[str] = []
    params: list = []
    if date_from:
        where_parts.append(f"{date_field} >= ?")
        params.append(date_from)
    if date_to:
        where_parts.append(f"{date_field} <= ?")
        params.append(date_to)
    where_sql = f" WHERE {' AND '.join(where_parts)}" if where_parts else ""

    rows = _query_to_json(
        f"""SELECT COALESCE(NULLIF(TRIM({row_field}), ''), '(не указано)') as row_label,
                   COALESCE(NULLIF(TRIM({col_field}), ''), '(не указано)') as col_label,
                   COUNT(*) as value
            FROM {table}{where_sql}
            GROUP BY row_label, col_label""",
        tuple(params),
    )

    # Собираем уникальные значения
    row_labels = sorted({r["row_label"] for r in rows})
    col_labels = sorted({r["col_label"] for r in rows})

    # Формируем матрицу кросс-таблицы
    matrix: list[dict[str, Any]] = []
    for r_label in row_labels:
        row_data: dict[str, Any] = {"row": r_label}
        for c_label in col_labels:
            row_data[c_label] = 0
        matrix.append(row_data)

    for r in rows:
        for row_data in matrix:
            if row_data["row"] == r["row_label"]:
                row_data[r["col_label"]] = r["value"]
                break

    return jsonify({
        "rows": matrix,
        "row_field": row_field,
        "col_field": col_field,
        "row_labels": row_labels,
        "col_labels": col_labels,
    })


# ---------------------------------------------------------------------------
# Отчёты
# ---------------------------------------------------------------------------

def _format_date_ru(value: str) -> str:
    """Преобразует дату ГГГГ-ММ-ДД в вид ДД.ММ.ГГГГ."""
    parts = str(value).split("-")
    if len(parts) != 3:
        return str(value)
    return f"{parts[2]}.{parts[1]}.{parts[0]}"


def _report_period_from_args() -> tuple[str, str, str] | tuple[None, None, str]:
    """Определяет период отчёта из параметров запроса.

    Поддерживаются два способа задания периода:
    - month=ГГГГ-ММ — один месяц (обратная совместимость);
    - date_from=ГГГГ-ММ-ДД и date_to=ГГГГ-ММ-ДД — произвольный период:
      один/несколько месяцев или конкретные даты.

    Возвращает (date_from, date_to, period_label).
    При ошибке возвращает (None, None, текст_ошибки).
    """
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    month = request.args.get("month", "").strip()

    if month:
        # Режим «один месяц» (для обратной совместимости)
        try:
            year, month_num = month.split("-")
            year = int(year)
            month_num = int(month_num)
            if not 1 <= month_num <= 12:
                raise ValueError
            if month_num == 12:
                last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
            else:
                last_day = datetime(year, month_num + 1, 1) - timedelta(days=1)
            date_from = f"{year:04d}-{month_num:02d}-01"
            date_to = f"{last_day.year:04d}-{last_day.month:02d}-{last_day.day:02d}"
        except ValueError:
            return None, None, "Неверный формат месяца. Используйте ГГГГ-ММ."
    elif not date_from or not date_to:
        return None, None, "Укажите период: дату «с» и дату «по»."
    else:
        try:
            datetime.strptime(date_from, "%Y-%m-%d")
            datetime.strptime(date_to, "%Y-%m-%d")
        except ValueError:
            return None, None, "Неверный формат даты. Используйте ГГГГ-ММ-ДД."
        if date_from > date_to:
            return None, None, "Дата «по» не может быть раньше даты «с»."

    period = f"{_format_date_ru(date_from)} - {_format_date_ru(date_to)}"
    return date_from, date_to, period


@app.route("/api/export", methods=["GET"])
def export():
    """Экспортирует записи с учётом текущих фильтров в CSV/XLSX/PDF."""
    table_type = request.args.get("table", "claims")
    fmt = request.args.get("format", "xlsx")
    if fmt not in ("csv", "xlsx", "pdf"):
        return jsonify({"error": "Неподдерживаемый формат экспорта."}), 400

    filters = {
        "search": request.args.get("search", ""),
        "date_from": request.args.get("date_from", ""),
        "date_to": request.args.get("date_to", ""),
    }
    for key, value in request.args.items():
        if key.startswith("eq_") or key.startswith("min_") or key.startswith("max_") or key.startswith("in_"):
            filters[key] = value

    where_sql, params = _apply_filters(table_type, filters)
    table = _table_name(table_type)
    rows = _query_to_json(f"SELECT * FROM {table}{where_sql}", tuple(params))

    title = "Претензии" if table_type == "claims" else "Операционный контроль"
    buffer = reports.export_data(rows, table_type, fmt, title)

    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    ext = {"csv": "csv", "xlsx": "xlsx", "pdf": "pdf"}[fmt]
    filename = f"otchet_{title.lower().replace(' ', '_')}_{now}.{ext}"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype={
            "csv": "text/csv; charset=utf-8",
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "pdf": "application/pdf",
        }[fmt],
    )


@app.route("/api/export-defects", methods=["GET"])
def export_defects():
    """Экспортирует производственный брак (претензии с признаком) в CSV/XLSX/PDF.

    Период задаётся параметрами date_from/date_to (или month=ГГГГ-ММ).
    PDF формируется по образцу сводного отчёта (со сводками).
    """
    fmt = request.args.get("format", "xlsx")
    if fmt not in ("csv", "xlsx", "pdf"):
        return jsonify({"error": "Неподдерживаемый формат экспорта."}), 400

    date_from, date_to, period = _report_period_from_args()
    if not date_from:
        return jsonify({"error": period}), 400

    rows = _query_to_json(
        """SELECT * FROM claims
           WHERE is_production_defect = 1
             AND date_received >= ? AND date_received <= ?
           ORDER BY date_received DESC""",
        (date_from, date_to))

    if fmt == "pdf":
        buffer = reports.build_defects_report_pdf(rows, period)
    else:
        buffer = reports.export_defects(rows, fmt, "Производственный брак")

    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"proizvodstvennyj_brak_{date_from}_{date_to}_{now}.{fmt}"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype={
            "csv": "text/csv; charset=utf-8",
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "pdf": "application/pdf",
        }[fmt],
    )


@app.route("/api/monthly-report", methods=["GET"])
def monthly_report():
    """Формирует сводный отчёт в PDF со сводками за выбранный период.

    Период задаётся параметрами date_from/date_to (или month=ГГГГ-ММ).
    """
    date_from, date_to, period = _report_period_from_args()
    if not date_from:
        return jsonify({"error": period}), 400

    claims_rows = _query_to_json(
        "SELECT * FROM claims WHERE date_received >= ? AND date_received <= ?",
        (date_from, date_to))
    control_rows = _query_to_json(
        "SELECT * FROM operational_control WHERE control_date >= ? AND control_date <= ?",
        (date_from, date_to))

    summary_claims = reports.monthly_report_summary(claims_rows, "claims")
    summary_control = reports.monthly_report_summary(control_rows, "control")

    buffer = reports.build_monthly_report_pdf(
        summary_claims, summary_control, period)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"period_report_{date_from}_{date_to}.pdf",
        mimetype="application/pdf",
    )


# ---------------------------------------------------------------------------
# Файловая система
# ---------------------------------------------------------------------------

@app.route("/api/claim-files/<int:claim_id>", methods=["GET"])
def claim_files(claim_id: int):
    """Возвращает список файлов в папке претензии."""
    rows = _query_to_json("SELECT folder_path FROM claims WHERE id = ?", (claim_id,))
    if not rows:
        return jsonify({"error": "Претензия не найдена."}), 404

    folder = Path(rows[0]["folder_path"] or "")
    files = []
    if folder.exists() and folder.is_dir():
        for path in sorted(folder.iterdir()):
            if path.is_file():
                files.append({
                    "name": path.name,
                    "size": path.stat().st_size,
                    "modified": datetime.fromtimestamp(path.stat().st_mtime)
                              .strftime("%d.%m.%Y %H:%M"),
                })
    return jsonify({"folder_path": str(folder), "files": files})


@app.route("/api/open-folder/<int:claim_id>", methods=["POST"])
def open_folder(claim_id: int):
    """Открывает папку претензии в проводнике Windows."""
    rows = _query_to_json("SELECT folder_path FROM claims WHERE id = ?", (claim_id,))
    if not rows:
        return jsonify({"error": "Претензия не найдена."}), 404

    folder = Path(rows[0]["folder_path"] or "")
    if not folder.exists():
        folder.mkdir(parents=True, exist_ok=True)

    try:
        if sys.platform == "win32":
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
        return jsonify({"message": "Папка открыта в проводнике."})
    except Exception as exc:
        return jsonify({"error": f"Не удалось открыть папку: {exc}"}), 500


@app.route("/api/upload/<int:claim_id>", methods=["POST"])
def upload_file(claim_id: int):
    """Прикрепляет файл к претензии (копирует в папку претензии)."""
    rows = _query_to_json("SELECT folder_path FROM claims WHERE id = ?", (claim_id,))
    if not rows:
        return jsonify({"error": "Претензия не найдена."}), 404

    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "Файл не выбран."}), 400

    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": f"Недопустимый тип файла: {ext or '(без расширения)'}"}), 400

    folder = Path(rows[0]["folder_path"] or "")
    folder.mkdir(parents=True, exist_ok=True)

    safe_name = secure_filename(file.filename)
    if not safe_name:
        safe_name = f"file_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    dest = folder / safe_name
    file.save(dest)

    log_change("claims", claim_id, "file_upload", f"Файл: {safe_name}")
    return jsonify({"message": "Файл прикреплён.", "filename": safe_name})


# ---------------------------------------------------------------------------
# Настройки, резервное копирование, журнал
# ---------------------------------------------------------------------------

@app.route("/api/settings", methods=["GET"])
def get_settings():
    """Возвращает текущие настройки приложения."""
    return jsonify({
        "claims_root": str(config.CLAIMS_ROOT),
        "db_path": str(config.DB_PATH),
        "backup_dir": str(config.BACKUP_DIR),
    })


@app.route("/api/backup", methods=["POST"])
def backup():
    """Создаёт резервную копию базы данных."""
    try:
        path = create_backup()
        if not path:
            return jsonify({"error": "База данных ещё не создана."}), 400
        return jsonify({"message": "Резервная копия создана.", "path": path})
    except Exception as exc:
        return jsonify({"error": f"Ошибка резервного копирования: {exc}"}), 500


@app.route("/api/logs", methods=["GET"])
def get_logs():
    """Возвращает журнал действий пользователей."""
    limit = min(max(int(request.args.get("limit", 200)), 1), 1000)
    rows = _query_to_json(
        "SELECT * FROM change_log ORDER BY id DESC LIMIT ?", (limit,))
    return jsonify({"data": rows})


# ---------------------------------------------------------------------------
# Сервер
# ---------------------------------------------------------------------------

def create_app() -> Flask:
    """Инициализирует БД и возвращает приложение."""
    init_db()
    # Резервная копия при каждом запуске
    try:
        create_backup()
    except Exception:
        pass  # Не критично при первом запуске
    return app


def main() -> None:
    """Точка входа: запускает сервер и открывает браузер."""
    host = "127.0.0.1"
    port = 5000

    init_db()
    try:
        create_backup()
    except Exception:
        pass

    # Автоматическое открытие браузера
    import threading
    import webbrowser

    def open_browser() -> None:
        webbrowser.open(f"http://{host}:{port}")

    threading.Timer(1.5, open_browser).start()

    try:
        app.run(host=host, port=port, debug=False)
    except OSError as exc:
        print(f"Ошибка запуска сервера: {exc}")
        print("Проверьте, что порт 5000 свободен, или измените порт в app.py")
        sys.exit(1)


if __name__ == "__main__":
    main()
"""
Модуль импорта данных из Excel-файлов (.xlsx).

Поддерживает два листа:
- Журнал претензий: лист «УТ_Рек»
- Операционный контроль: лист «База»
"""
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

import config
from database import get_connection, log_change

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

# Сопоставление столбцов листа «УТ_Рек» с полями таблицы claims
CLAIM_COLUMN_MAP = {
    "№ п/п": "claim_number",
    "Дата обращения": "date_received",
    "Способ обращения": "contact_method",
    "Заявитель": "applicant",
    "Адрес заявителя": "address",
    "Контакты": "contacts",
    "Наименование товара": "product_name",
    "Дефект": "defect",
    "Количество": "quantity",
    "Статус": "status",
    "Дата решения": "decision_date",
    "Дата отгрузки": "shipment_date",
    "Возврат": "return_date",
}

# Сопоставление столбцов листа «База» с полями таблицы operational_control
CONTROL_COLUMN_MAP = {
    "Дата": "control_date",
    "Партия": "batch",
    "Прессовщик": "press_operator",
    "Форма": "mold",
    "Мастер смены": "shift_master",
    "Дефект": "defect",
    "Количество": "quantity",
    "Время проверки": "check_time",
}


class ImportError(Exception):
    """Пользовательская ошибка импорта с понятным сообщением."""


def _normalize_header(value: Any) -> str:
    """Приводит заголовок столбца к строке и убирает лишние пробелы."""
    if value is None:
        return ""
    return str(value).strip().replace("\n", " ").replace("\r", "")


def _parse_date(value: Any) -> str | None:
    """Преобразует значение Excel в строку даты ISO (YYYY-MM-DD)."""
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    # openpyxl возвращает datetime для ячеек с форматом даты
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    # Попытка разбора строковых дат
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    # Если это число Excel (серийный номер даты)
    try:
        serial = float(text)
        if 1 <= serial <= 2958465:
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=serial)).strftime("%Y-%m-%d")
    except ValueError:
        pass
    return text  # Возвращаем как есть, если не удалось распознать


def _parse_time(value: Any) -> str | None:
    """Преобразует значение во время (HH:MM)."""
    if value is None or str(value).strip() == "":
        return None
    if hasattr(value, "strftime"):
        return value.strftime("%H:%M")
    text = str(value).strip()
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%H:%M")
        except ValueError:
            continue
    return text


def _parse_int(value: Any) -> int | None:
    """Безопасное преобразование в целое число."""
    if value is None or str(value).strip() == "":
        return None
    try:
        return int(float(str(value).replace(" ", "").replace(",", ".")))
    except (ValueError, TypeError):
        return None


def load_workbook_safe(path: str | Path) -> Any:
    """Загружает книгу Excel (.xlsx или .xls) с понятным сообщением об ошибке."""
    path_str = str(path)
    try:
        if path_str.lower().endswith(".xls"):
            if not HAS_XLRD:
                raise ImportError(
                    "Для чтения файлов .xls требуется библиотека xlrd. "
                    "Установите её: pip install xlrd"
                )
            return xlrd.open_workbook(path_str)
        return load_workbook(path_str, data_only=True)
    except FileNotFoundError:
        raise ImportError(f"Файл не найден: {path}")
    except ImportError:
        raise
    except Exception as exc:
        raise ImportError(f"Не удалось открыть Excel-файл. Возможно, файл повреждён "
                          f"или имеет неверный формат. Подробности: {exc}")


def _read_rows(ws, column_map: dict[str, str]) -> list[dict[str, Any]]:
    """Читает строки с листа, сопоставляя заголовки с полями БД.

    Поддерживает листы openpyxl (.xlsx) и xlrd (.xls).
    """
    # Определяем тип листа
    is_xls = hasattr(ws, "nrows")  # xlrd sheet имеет nrows

    def get_cell_value(row, col):
        if is_xls:
            cell = ws.cell(row, col)
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    return xlrd.xldate_as_datetime(cell.value, ws.book.datemode)
                except Exception:
                    return cell.value
            return cell.value
        return row[col].value

    header_row = None

    if is_xls:
        # xlrd: строки 0-based, столбцы 0-based
        for row_idx in range(ws.nrows):
            values = [_normalize_header(ws.cell(row_idx, col).value)
                      for col in range(ws.ncols)]
            if any(key in values for key in column_map):
                header_row = row_idx
                break
    else:
        # openpyxl: строки 1-based
        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            values = [_normalize_header(cell.value) for cell in row]
            if any(key in values for key in column_map):
                header_row = row_idx
                break

    if header_row is None:
        expected = ", ".join(column_map.keys())
        raise ImportError(f"Не найден лист с ожидаемыми столбцами. Ожидаемые столбцы: {expected}")

    if is_xls:
        headers = [_normalize_header(ws.cell(header_row, col).value)
                   for col in range(ws.ncols)]
    else:
        headers = [_normalize_header(cell.value) for cell in ws[header_row]]

    mapping: dict[int, str] = {}
    for idx, header in enumerate(headers):
        if header in column_map:
            mapping[idx] = column_map[header]

    if not mapping:
        raise ImportError("Не удалось сопоставить ни одного столбца с полями базы данных.")

    rows: list[dict[str, Any]] = []
    if is_xls:
        for row_idx in range(header_row + 1, ws.nrows):
            record: dict[str, Any] = {}
            empty = True
            for col_idx, field in mapping.items():
                value = get_cell_value(row_idx, col_idx)
                if value is not None and str(value).strip() != "":
                    empty = False
                record[field] = value
            if not empty:
                rows.append(record)
    else:
        for row in ws.iter_rows(min_row=header_row + 1):
            record = {}
            empty = True
            for col_idx, field in mapping.items():
                value = row[col_idx].value
                if value is not None and str(value).strip() != "":
                    empty = False
                record[field] = value
            if not empty:
                rows.append(record)

    return rows


def _normalize_claim(record: dict[str, Any]) -> dict[str, Any]:
    """Приводит запись претензии к типам БД."""
    normalized = {
        "claim_number": _parse_int(record.get("claim_number")),
        "date_received": _parse_date(record.get("date_received")),
        "contact_method": str(record["contact_method"]).strip() if record.get("contact_method") else None,
        "applicant": str(record["applicant"]).strip() if record.get("applicant") else None,
        "address": str(record["address"]).strip() if record.get("address") else None,
        "contacts": str(record["contacts"]).strip() if record.get("contacts") else None,
        "product_name": str(record["product_name"]).strip() if record.get("product_name") else None,
        "defect": str(record["defect"]).strip() if record.get("defect") else None,
        "quantity": _parse_int(record.get("quantity")),
        "status": str(record["status"]).strip() if record.get("status") else None,
        "decision_date": _parse_date(record.get("decision_date")),
        "shipment_date": _parse_date(record.get("shipment_date")),
        "return_date": _parse_date(record.get("return_date")),
    }
    if normalized["claim_number"] is None:
        raise ImportError("Запись пропущена: отсутствует номер претензии (№ п/п).")
    return normalized


def _normalize_control(record: dict[str, Any]) -> dict[str, Any]:
    """Приводит запись операционного контроля к типам БД."""
    return {
        "control_date": _parse_date(record.get("control_date")),
        "batch": str(record["batch"]).strip() if record.get("batch") else None,
        "press_operator": str(record["press_operator"]).strip() if record.get("press_operator") else None,
        "mold": str(record["mold"]).strip() if record.get("mold") else None,
        "shift_master": str(record["shift_master"]).strip() if record.get("shift_master") else None,
        "defect": str(record["defect"]).strip() if record.get("defect") else None,
        "quantity": _parse_int(record.get("quantity")),
        "check_time": _parse_time(record.get("check_time")),
    }


def _build_folder_path(claim_number: int, date_received: str | None) -> str:
    """Формирует путь к папке претензии по шаблону."""
    date_part = date_received or "bez_daty"
    folder_name = f"pretensiya_{claim_number}_{date_part}"
    folder_name_safe = folder_name.replace(":", "_").replace("\\", "_").replace("/", "_")
    return str(config.CLAIMS_ROOT / folder_name_safe)


def _get_sheet_names(wb: Any) -> list[str]:
    """Возвращает имена листов для openpyxl и xlrd."""
    if hasattr(wb, "sheetnames"):
        return wb.sheetnames
    return wb.sheet_names()


def _get_sheet(wb: Any, name: str):
    """Возвращает лист по имени для openpyxl и xlrd."""
    if hasattr(wb, "sheetnames"):
        return wb[name]
    return wb.sheet_by_name(name)


def import_claims(path: str | Path) -> dict[str, int]:
    """Импортирует претензии из листа «УТ_Рек».

    Обновляет существующие записи по номеру претензии, добавляет новые.
    Возвращает статистику: {added, updated, skipped}.
    """
    wb = load_workbook_safe(path)
    sheet_names = _get_sheet_names(wb)
    if "УТ_Рек" not in sheet_names:
        sheets = ", ".join(sheet_names)
        raise ImportError(f"В файле отсутствует лист «УТ_Рек». Доступны листы: {sheets}")

    ws = _get_sheet(wb, "УТ_Рек")
    rows = _read_rows(ws, CLAIM_COLUMN_MAP)
    if not rows:
        return {"added": 0, "updated": 0, "skipped": 0}

    added = updated = skipped = 0
    with get_connection() as conn:
        for raw in rows:
            try:
                rec = _normalize_claim(raw)
            except ImportError:
                skipped += 1
                continue

            existing = conn.execute(
                "SELECT id FROM claims WHERE claim_number = ?", (rec["claim_number"],)
            ).fetchone()

            if existing:
                conn.execute(
                    """UPDATE claims SET
                        date_received = ?, contact_method = ?, applicant = ?,
                        address = ?, contacts = ?, product_name = ?, defect = ?,
                        quantity = ?, status = ?, decision_date = ?,
                        shipment_date = ?, return_date = ?
                       WHERE id = ?""",
                    (rec["date_received"], rec["contact_method"], rec["applicant"],
                     rec["address"], rec["contacts"], rec["product_name"], rec["defect"],
                     rec["quantity"], rec["status"], rec["decision_date"],
                     rec["shipment_date"], rec["return_date"], existing["id"]),
                )
                log_change("claims", existing["id"], "update",
                           f"Импорт: претензия {rec['claim_number']}", conn=conn)
                updated += 1
            else:
                folder_path = _build_folder_path(rec["claim_number"], rec["date_received"])
                cur = conn.execute(
                    """INSERT INTO claims
                       (claim_number, date_received, contact_method, applicant,
                        address, contacts, product_name, defect, quantity,
                        status, decision_date, shipment_date, return_date, folder_path)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (rec["claim_number"], rec["date_received"], rec["contact_method"],
                     rec["applicant"], rec["address"], rec["contacts"], rec["product_name"],
                     rec["defect"], rec["quantity"], rec["status"], rec["decision_date"],
                     rec["shipment_date"], rec["return_date"], folder_path),
                )
                log_change("claims", cur.lastrowid, "insert",
                           f"Импорт: претензия {rec['claim_number']}", conn=conn)
                added += 1

    return {"added": added, "updated": updated, "skipped": skipped}


def import_operational_control(path: str | Path) -> dict[str, int]:
    """Импортирует операционный контроль из листа «База».

    Обновляет существующие записи по номеру партии, добавляет новые.
    """
    wb = load_workbook_safe(path)
    sheet_names = _get_sheet_names(wb)
    if "База" not in sheet_names:
        sheets = ", ".join(sheet_names)
        raise ImportError(f"В файле отсутствует лист «База». Доступны листы: {sheets}")
    ws = _get_sheet(wb, "База")
    rows = _read_rows(ws, CONTROL_COLUMN_MAP)
    if not rows:
        return {"added": 0, "updated": 0, "skipped": 0}

    added = updated = skipped = 0
    with get_connection() as conn:
        for raw in rows:
            rec = _normalize_control(raw)
            if not rec["batch"]:
                skipped += 1
                continue

            existing = conn.execute(
                "SELECT id FROM operational_control WHERE batch = ?", (rec["batch"],)
            ).fetchone()

            if existing:
                conn.execute(
                    """UPDATE operational_control SET
                        control_date = ?, press_operator = ?, mold = ?,
                        shift_master = ?, defect = ?, quantity = ?, check_time = ?
                       WHERE id = ?""",
                    (rec["control_date"], rec["press_operator"], rec["mold"],
                     rec["shift_master"], rec["defect"], rec["quantity"], rec["check_time"],
                     existing["id"]),
                )
                log_change("operational_control", existing["id"], "update",
                           f"Импорт: партия {rec['batch']}", conn=conn)
                updated += 1
            else:
                cur = conn.execute(
                    """INSERT INTO operational_control
                       (control_date, batch, press_operator, mold, shift_master,
                        defect, quantity, check_time)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (rec["control_date"], rec["batch"], rec["press_operator"],
                     rec["mold"], rec["shift_master"], rec["defect"],
                     rec["quantity"], rec["check_time"]),
                )
                log_change("operational_control", cur.lastrowid, "insert",
                           f"Импорт: партия {rec['batch']}", conn=conn)
                added += 1

    return {"added": added, "updated": updated, "skipped": skipped}
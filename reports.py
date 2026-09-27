"""
Модуль формирования и экспорта отчётов.

Поддерживаемые форматы: CSV, XLSX, PDF.
Также формируется ежемесячная сводка по претензиям и операционному контролю.
"""
import csv
import io
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

# Регистрация шрифтов с поддержкой кириллицы (Arial из системных шрифтов Windows)
_FONT_DIR = r"C:\Windows\Fonts"
try:
    pdfmetrics.registerFont(TTFont("Arial", f"{_FONT_DIR}\\arial.ttf"))
    pdfmetrics.registerFont(TTFont("Arial-Bold", f"{_FONT_DIR}\\arialbd.ttf"))
    pdfmetrics.registerFontFamily("Arial", normal="Arial", bold="Arial-Bold",
                                  italic="Arial", boldItalic="Arial-Bold")
    FONT_NAME = "Arial"
    FONT_BOLD = "Arial-Bold"
except Exception:
    # Если системные шрифты недоступны, используем стандартные (без кириллицы)
    FONT_NAME = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"

CLAIM_FIELDS = [
    ("claim_number", "№ п/п"),
    ("date_received", "Дата обращения"),
    ("contact_method", "Способ обращения"),
    ("applicant", "Заявитель"),
    ("address", "Адрес заявителя"),
    ("contacts", "Контакты"),
    ("product_name", "Наименование товара"),
    ("defect", "Дефект"),
    ("quantity", "Количество"),
    ("status", "Статус"),
    ("decision_date", "Дата решения"),
    ("shipment_date", "Дата отгрузки"),
    ("return_date", "Возврат"),
]

CONTROL_FIELDS = [
    ("control_date", "Дата"),
    ("batch", "Партия"),
    ("press_operator", "Прессовщик"),
    ("mold", "Форма"),
    ("shift_master", "Мастер смены"),
    ("defect", "Дефект"),
    ("quantity", "Количество"),
    ("check_time", "Время проверки"),
]

# Поля для отчёта по производственному браку (только претензии с признаком)
DEFECT_FIELDS = [
    ("claim_number", "№ п/п"),
    ("date_received", "Дата обращения"),
    ("contact_method", "Способ обращения"),
    ("applicant", "Заявитель"),
    ("address", "Адрес заявителя"),
    ("contacts", "Контакты"),
    ("product_name", "Наименование товара"),
    ("defect", "Дефект"),
    ("quantity", "Количество"),
    ("status", "Статус"),
    ("decision_date", "Дата решения"),
    ("shipment_date", "Дата отгрузки"),
    ("return_date", "Возврат"),
]

STATUS_COLORS = {
    "Согласовано": "2E7D32",
    "Отказ": "C62828",
    "Ждем ответ": "F9A825",
    "Ждём ответ": "F9A825",
}


def _rows_to_csv(data: list[dict[str, Any]], fields: list[tuple[str, str]]) -> BytesIO:
    """Преобразует записи в CSV (UTF-8 с BOM для Excel)."""
    buffer = BytesIO()
    # Записываем BOM для корректного открытия в Excel
    buffer.write("\ufeff".encode("utf-8"))
    # csv.writer требует текстовый поток, поэтому используем StringIO
    text_buffer = io.StringIO()
    writer = csv.writer(text_buffer)
    writer.writerow([label for _, label in fields])
    for row in data:
        writer.writerow([row.get(key, "") for key, _ in fields])
    buffer.write(text_buffer.getvalue().encode("utf-8"))
    buffer.seek(0)
    return buffer


def _rows_to_xlsx(data: list[dict[str, Any]], fields: list[tuple[str, str]]) -> BytesIO:
    """Преобразует записи в XLSX с форматированием заголовков."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1976D2", end_color="1976D2", fill_type="solid")

    # Заголовки
    for col_idx, (_, label) in enumerate(fields, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Данные
    for row_idx, row in enumerate(data, start=2):
        for col_idx, (key, _) in enumerate(fields, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row.get(key, ""))

    # Автоширина столбцов
    for col_idx in range(1, len(fields) + 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                value = str(cell.value or "")
                max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

    ws.freeze_panes = "A2"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def _make_pdf_table_style() -> TableStyle:
    """Единый стиль таблиц для PDF-отчётов."""
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1976D2")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("FONTNAME", (0, 1), (-1, -1), FONT_NAME),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BDBDBD")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ])


def _rows_to_pdf(data: list[dict[str, Any]], fields: list[tuple[str, str]],
                 title: str) -> BytesIO:
    """Преобразует записи в PDF-файл.

    Таблица растягивается на всю ширину ландшафтной A4, при необходимости
    уменьшается шрифт, чтобы все колонки поместились (не обрезались).
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=10 * mm, leftMargin=10 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleRu", parent=styles["Title"], fontSize=14, spaceAfter=8,
        fontName=FONT_BOLD
    )
    date_style = ParagraphStyle("DateRu", parent=styles["Normal"], fontSize=8,
                                fontName=FONT_NAME)

    elements = [
        Paragraph(title, title_style),
        Paragraph(f"Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}", date_style),
        Spacer(1, 6 * mm),
    ]

    # Адаптивный размер шрифта в зависимости от числа колонок
    col_count = len(fields)
    if col_count > 10:
        font_size = 5.5
        header_font_size = 6
        cell_padding = 1.5
    elif col_count > 8:
        font_size = 6.5
        header_font_size = 7
        cell_padding = 2
    else:
        font_size = 7.5
        header_font_size = 8
        cell_padding = 3

    # Стиль для ячеек с переносом длинных текстов
    cell_style = ParagraphStyle(
        "CellRu", fontName=FONT_NAME, fontSize=font_size, leading=font_size + 1
    )
    header_cell_style = ParagraphStyle(
        "HeaderCellRu", fontName=FONT_BOLD, fontSize=header_font_size,
        leading=header_font_size + 1, textColor=colors.white
    )

    header = [Paragraph(label, header_cell_style) for _, label in fields]
    table_data = [header]
    for row in data:
        table_data.append([
            Paragraph(str(row.get(key, "")), cell_style) for key, _ in fields
        ])

    # Ширина таблицы = ширина страницы минус поля
    usable_width = landscape(A4)[0] - 20 * mm
    col_width = usable_width / col_count

    table = Table(
        table_data,
        repeatRows=1,
        colWidths=[col_width] * col_count,
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1976D2")),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDBDBD")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), cell_padding),
        ("RIGHTPADDING", (0, 0), (-1, -1), cell_padding),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer


def export_data(data: list[dict[str, Any]], table_type: str,
                fmt: str, title: str = "Отчет") -> BytesIO:
    """Экспортирует список записей в указанном формате.

    table_type: "claims" или "control"
    fmt: "csv", "xlsx" или "pdf"
    """
    fields = CLAIM_FIELDS if table_type == "claims" else CONTROL_FIELDS
    if fmt == "csv":
        return _rows_to_csv(data, fields)
    if fmt == "xlsx":
        return _rows_to_xlsx(data, fields)
    if fmt == "pdf":
        return _rows_to_pdf(data, fields, title)
    raise ValueError(f"Неподдерживаемый формат: {fmt}")


def export_defects(data: list[dict[str, Any]], fmt: str,
                   title: str = "Производственный брак") -> BytesIO:
    """Экспортирует записи производственного брака (претензии с признаком).

    fmt: "csv", "xlsx" или "pdf"
    """
    if fmt == "csv":
        return _rows_to_csv(data, DEFECT_FIELDS)
    if fmt == "xlsx":
        return _rows_to_xlsx(data, DEFECT_FIELDS)
    if fmt == "pdf":
        return _rows_to_pdf(data, DEFECT_FIELDS, title)
    raise ValueError(f"Неподдерживаемый формат: {fmt}")


def monthly_report_summary(rows: list[dict[str, Any]], table_type: str) -> dict[str, Any]:
    """Строит сводку по записям за период (для ежемесячного отчёта).

    Возвращает словарь со статистикой:
    - total: общее количество
    - by_status: распределение по статусам (для претензий)
    - top_products: топ-5 товаров (для претензий)
    - top_defects: топ-5 дефектов
    - total_quantity: суммарное количество брака (для ОК)
    - by_operator / by_mold: статистика по прессовщикам и формам (для ОК)
    """
    total = len(rows)
    result: dict[str, Any] = {"total": total}

    def count_by(key: str, limit: int | None = None) -> list[tuple[str, int]]:
        counter: dict[str, int] = {}
        for row in rows:
            value = str(row.get(key) or "").strip()
            if not value:
                value = "(не указано)"
            counter[value] = counter.get(value, 0) + 1
        items = sorted(counter.items(), key=lambda x: x[1], reverse=True)
        if limit:
            items = items[:limit]
        return items

    if table_type == "claims":
        result["by_status"] = count_by("status")
        result["top_products"] = count_by("product_name", 5)
        result["top_defects"] = count_by("defect", 5)
    else:
        result["by_defect"] = count_by("defect", 5)
        result["by_operator"] = count_by("press_operator")
        result["by_mold"] = count_by("mold")
        result["by_master"] = count_by("shift_master")
        result["total_quantity"] = sum(
            int(row.get("quantity") or 0) for row in rows
        )

    return result


def build_defects_report_pdf(rows: list[dict[str, Any]], period: str) -> BytesIO:
    """Формирует отчёт по производственному браку в PDF со сводками.

    По образцу ежемесячного отчёта: заголовок, период, сводки
    (общее количество, по статусам, топ товаров, топ дефектов)
    и детальная таблица всех записей.
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=15 * mm, leftMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        "NormalRu", parent=styles["Normal"], fontName=FONT_NAME, fontSize=10
    )
    title_style = ParagraphStyle(
        "TitleRu", parent=styles["Title"], fontSize=16, spaceAfter=6,
        fontName=FONT_BOLD
    )
    section_style = ParagraphStyle(
        "SectionRu", parent=styles["Heading2"], fontSize=12,
        spaceBefore=12, spaceAfter=4, textColor=colors.HexColor("#1976D2"),
        fontName=FONT_BOLD
    )

    elements = [
        Paragraph("Отчёт по производственному браку", title_style),
        Paragraph(f"Период: {period}", normal_style),
        Paragraph(f"Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                  normal_style),
    ]

    # Сводки
    summary = monthly_report_summary(rows, "claims")
    elements.append(Paragraph(
        f"Общее количество записей: <b>{summary.get('total', 0)}</b>",
        normal_style
    ))

    if summary.get("by_status"):
        elements.append(Paragraph("Распределение по статусам:", section_style))
        table = Table([["Статус", "Количество"]] + list(summary["by_status"]))
        table.setStyle(_make_pdf_table_style())
        elements.append(table)

    if summary.get("top_products"):
        elements.append(Paragraph("Топ-5 товаров по количеству претензий:", section_style))
        table = Table([["Товар", "Количество"]] + list(summary["top_products"]))
        table.setStyle(_make_pdf_table_style())
        elements.append(table)

    if summary.get("top_defects"):
        elements.append(Paragraph("Топ-5 дефектов:", section_style))
        table = Table([["Дефект", "Количество"]] + list(summary["top_defects"]))
        table.setStyle(_make_pdf_table_style())
        elements.append(table)

    # Детальная таблица
    elements.append(Paragraph("Детализация по претензиям:", section_style))
    if not rows:
        elements.append(Paragraph("Нет данных за выбранный период.", normal_style))
    else:
        # Адаптивный размер шрифта
        col_count = len(DEFECT_FIELDS)
        font_size = 5.5 if col_count > 10 else 6.5
        header_font_size = 6 if col_count > 10 else 7
        cell_padding = 1.5 if col_count > 10 else 2

        cell_style = ParagraphStyle(
            "CellRu", fontName=FONT_NAME, fontSize=font_size, leading=font_size + 1
        )
        header_cell_style = ParagraphStyle(
            "HeaderCellRu", fontName=FONT_BOLD, fontSize=header_font_size,
            leading=header_font_size + 1, textColor=colors.white
        )

        header = [Paragraph(label, header_cell_style) for _, label in DEFECT_FIELDS]
        table_data = [header]
        for row in rows:
            table_data.append([
                Paragraph(str(row.get(key, "")), cell_style) for key, _ in DEFECT_FIELDS
            ])

        usable_width = A4[0] - 30 * mm
        col_width = usable_width / col_count
        table = Table(table_data, repeatRows=1, colWidths=[col_width] * col_count)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1976D2")),
            ("FONTSIZE", (0, 0), (-1, -1), font_size),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDBDBD")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), cell_padding),
            ("RIGHTPADDING", (0, 0), (-1, -1), cell_padding),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer


def build_monthly_report_pdf(summary_claims: dict[str, Any] | None,
                             summary_control: dict[str, Any] | None,
                             period: str) -> BytesIO:
    """Формирует ежемесячный отчёт в PDF со сводками."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=15 * mm, leftMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        "NormalRu", parent=styles["Normal"], fontName=FONT_NAME, fontSize=10
    )
    title_style = ParagraphStyle(
        "TitleRu", parent=styles["Title"], fontSize=16, spaceAfter=6,
        fontName=FONT_BOLD
    )
    section_style = ParagraphStyle(
        "SectionRu", parent=styles["Heading2"], fontSize=12,
        spaceBefore=12, spaceAfter=4, textColor=colors.HexColor("#1976D2"),
        fontName=FONT_BOLD
    )

    elements = [
        Paragraph("Сводный отчёт по качеству", title_style),
        Paragraph(f"Период: {period}", normal_style),
        Paragraph(f"Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                  normal_style),
    ]

    def add_section(name: str, summary: dict[str, Any]) -> None:
        elements.append(Paragraph(name, section_style))
        if summary is None:
            elements.append(Paragraph("Нет данных за выбранный период.", normal_style))
            return

        elements.append(Paragraph(
            f"Общее количество записей: <b>{summary.get('total', 0)}</b>",
            normal_style
        ))

        if "by_status" in summary and summary["by_status"]:
            elements.append(Paragraph("Распределение по статусам:", normal_style))
            table = Table([["Статус", "Количество"]] + list(summary["by_status"]))
            table.setStyle(_make_pdf_table_style())
            elements.append(table)

        if "top_products" in summary and summary["top_products"]:
            elements.append(Paragraph("Топ-5 товаров по количеству претензий:", normal_style))
            table = Table([["Товар", "Количество"]] + list(summary["top_products"]))
            table.setStyle(_make_pdf_table_style())
            elements.append(table)

        if "top_defects" in summary and summary["top_defects"]:
            elements.append(Paragraph("Топ-5 дефектов:", normal_style))
            table = Table([["Дефект", "Количество"]] + list(summary["top_defects"]))
            table.setStyle(_make_pdf_table_style())
            elements.append(table)

        if "total_quantity" in summary:
            elements.append(Paragraph(
                f"Суммарное количество брака: <b>{summary.get('total_quantity', 0)}</b>",
                normal_style
            ))

        if "by_operator" in summary and summary["by_operator"]:
            elements.append(Paragraph("Статистика по прессовщикам:", normal_style))
            table = Table([["Прессовщик", "Количество"]] + list(summary["by_operator"]))
            table.setStyle(_make_pdf_table_style())
            elements.append(table)

        if "by_mold" in summary and summary["by_mold"]:
            elements.append(Paragraph("Статистика по формам:", normal_style))
            table = Table([["Форма", "Количество"]] + list(summary["by_mold"]))
            table.setStyle(_make_pdf_table_style())
            elements.append(table)

    add_section("Рекламации (претензии заказчиков)", summary_claims)
    add_section("Операционный контроль", summary_control)

    doc.build(elements)
    buffer.seek(0)
    return buffer
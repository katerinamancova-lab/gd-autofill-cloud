"""Workbook inspection, safe filling and report sheet generation."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from config import PROTECTED_COLUMN_MARKERS, REPORT_SHEETS

YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill(fill_type="solid", fgColor="17365D")
HEADER_FONT = Font(color="FFFFFF", bold=True)


@dataclass
class WorkbookLayout:
    sheet: Worksheet
    header_row: int
    headers: list[str]
    name_column: int
    data_rows: list[int]


def _text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _name_score(header: str) -> int:
    normalized = header.lower()
    exact = {
        "наименование": 100,
        "название": 95,
        "наименование товара": 110,
        "название товара": 105,
        "name": 90,
        "product name": 105,
        "модель": 80,
    }
    return exact.get(normalized, 60 if "наимен" in normalized or "назван" in normalized else 0)


def inspect_workbook(file_bytes: bytes) -> tuple[Any, WorkbookLayout]:
    workbook = load_workbook(io.BytesIO(file_bytes))
    candidates: list[WorkbookLayout] = []
    for sheet in workbook.worksheets:
        if sheet.title in REPORT_SHEETS:
            continue
        for row in range(1, min(sheet.max_row, 30) + 1):
            headers = [_text(sheet.cell(row, col).value) for col in range(1, sheet.max_column + 1)]
            if sum(bool(header) for header in headers) < 2:
                continue
            scored = [(index + 1, _name_score(header)) for index, header in enumerate(headers)]
            name_column, score = max(scored, key=lambda item: item[1])
            if score <= 0:
                continue
            data_rows = [
                index
                for index in range(row + 1, sheet.max_row + 1)
                if _text(sheet.cell(index, name_column).value)
            ]
            if data_rows:
                candidates.append(WorkbookLayout(sheet, row, headers, name_column, data_rows))
    if not candidates:
        raise ValueError(
            "Не найден лист с колонкой названия товара. Ожидается заголовок вроде "
            "«Наименование товара», «Название» или «Модель»."
        )
    layout = max(candidates, key=lambda item: len(item.data_rows))
    return workbook, layout


def is_protected(header: str) -> bool:
    normalized = header.strip().lower()
    return any(marker == normalized or marker in normalized for marker in PROTECTED_COLUMN_MARKERS)


def writable_columns(layout: WorkbookLayout, row: int) -> list[str]:
    output = []
    for index, header in enumerate(layout.headers, start=1):
        if not header or index == layout.name_column or is_protected(header):
            continue
        if layout.sheet.cell(row, index).value in (None, ""):
            output.append(header)
    return output


def write_values(layout: WorkbookLayout, row: int, values: dict[str, dict[str, str]]) -> int:
    header_to_column = {
        header: index for index, header in enumerate(layout.headers, start=1) if header
    }
    changed = 0
    for header, item in values.items():
        column = header_to_column.get(header)
        if not column or is_protected(header):
            continue
        cell = layout.sheet.cell(row, column)
        if cell.value not in (None, ""):
            continue
        cell.value = item["value"]
        cell.fill = YELLOW_FILL
        changed += 1
    return changed


def _replace_sheet(workbook: Any, name: str) -> Worksheet:
    if name in workbook.sheetnames:
        del workbook[name]
    return workbook.create_sheet(name)


def _write_table(sheet: Worksheet, headers: list[str], rows: list[list[Any]]) -> None:
    sheet.append(headers)
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True)
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in sheet.columns:
        letter = column[0].column_letter
        width = min(70, max(12, max(len(_text(cell.value)) for cell in column) + 2))
        sheet.column_dimensions[letter].width = width
        for cell in column:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def add_report_sheets(
    workbook: Any,
    report_rows: list[list[Any]],
    review_rows: list[list[Any]],
    source_rows: list[list[Any]],
) -> None:
    report = _replace_sheet(workbook, "Отчёт")
    _write_table(
        report,
        [
            "Строка товара",
            "Название товара",
            "Категория",
            "Уверенность категории",
            "Полей заполнено",
            "Источников найдено",
            "Сайтов открыто",
            "Сайтов пропущено",
            "Полей на проверку",
        ],
        report_rows,
    )
    review = _replace_sheet(workbook, "Проверить")
    _write_table(
        review,
        ["Строка товара", "Название товара", "Поле", "Причина", "Спорные значения"],
        review_rows,
    )
    sources = _replace_sheet(workbook, "Источники")
    _write_table(
        sources,
        [
            "Строка товара",
            "Название товара",
            "URL источника",
            "Провайдер поиска",
            "Статус",
            "Использован",
            "Ошибка",
        ],
        source_rows,
    )


def workbook_bytes(workbook: Any) -> bytes:
    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()

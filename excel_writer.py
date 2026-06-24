import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

YELLOW = PatternFill(fill_type="solid", fgColor="FFFF00")


def header_map(ws):
    return {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}


def set_if_col(ws, hmap, row, header, value):
    if value is None or value == "":
        return 0
    col = hmap.get(header)
    if not col:
        return 0
    if any(x in header.lower() for x in ["uid", "уид", "активность", "розничная цена"]):
        return 0
    cell = ws.cell(row, col)
    if str(cell.value or "").strip() == str(value or "").strip():
        return 0
    cell.value = value
    cell.fill = YELLOW
    return 1


def prepare_workbook(uploaded_file):
    wb = load_workbook(uploaded_file)
    ws = wb.active

    for s in ["Отчет", "Проверить", "Источники"]:
        if s in wb.sheetnames:
            del wb[s]

    report = wb.create_sheet("Отчет")
    report.append(["Показатель", "Значение"])

    check = wb.create_sheet("Проверить")
    check.append(["Строка", "Товар", "Поле", "Значение", "Комментарий"])

    sources = wb.create_sheet("Источники")
    sources.append(["Строка", "Товар", "URL", "Статус"])

    return wb, ws, report, check, sources


def export_workbook(wb):
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

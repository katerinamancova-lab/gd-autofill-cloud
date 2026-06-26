from io import BytesIO

import pytest
from openpyxl import Workbook

from category_rules import detect_category
from excel_writer import inspect_workbook, writable_columns, write_values


CATEGORY_FILES = {
    "Гольфкары.xlsx": "Гольфкары",
    "Электромотоциклы.xlsx": "Электромотоциклы",
    "Мотобуксировщики.xlsx": "Мотобуксировщики",
    "Снегоуборщики.xlsx": "Снегоуборщики",
    "Снегоходы.xlsx": "Снегоходы",
    "Лодки ПВХ.xlsx": "Лодки ПВХ",
    "Лодочные моторы.xlsx": "Лодочные моторы",
    "Дорожные мотоциклы.xlsx": "Дорожные мотоциклы",
    "Внедорожные мотоциклы.xlsx": "Внедорожные мотоциклы",
    "Импорт на ГД квадроциклы.xlsx": "Квадроциклы",
}


def workbook_bytes(product_name: str) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Квадроциклы"
    sheet.append(
        [
            "Заполнить в админке по необходимости:",
            "Наименование элемента",
            "Бренд [BRAND]",
            "Мощность, л.с. [POWER_HP]",
            "Мощность, л.с. [POWER_HP]",
        ]
    )
    sheet.append(["Сортировка", product_name, "Бренд", None, None])
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def test_all_exact_filename_categories():
    for filename, expected in CATEGORY_FILES.items():
        category, _confidence = detect_category(
            "Тестовая модель", ["Наименование элемента"], "Ошибочный лист", filename
        )
        assert category == expected


def test_empty_single_product_template_has_clear_error():
    with pytest.raises(ValueError, match="B2"):
        inspect_workbook(workbook_bytes("Квадроцикл БРЕНД Модель"))


def test_single_product_template_protects_column_a_and_fills_duplicates():
    workbook, layout = inspect_workbook(workbook_bytes("Tohatsu MFS50WETL"))
    assert layout.name_column == 2
    assert 1 in layout.excluded_columns
    assert "Заполнить в админке по необходимости:" not in writable_columns(layout, 2)
    changed = write_values(
        layout,
        2,
        {"Мощность, л.с. [POWER_HP]": {"value": "50", "source": "https://example.com"}},
    )
    assert changed == 2
    assert layout.sheet["A2"].value == "Сортировка"
    assert layout.sheet["D2"].value == "50"
    assert layout.sheet["E2"].value == "50"

"""Automatic product category detection."""

from __future__ import annotations

import re
from collections.abc import Iterable


CATEGORY_RULES: dict[str, tuple[str, ...]] = {
    "Лодочные моторы": (
        "лодочн",
        "подвесной мотор",
        "outboard",
        "bf",
        "f ",
        "df",
        "mfs",
        "tohatsu",
    ),
    "Лодки ПВХ": ("лодка пвх", "надувная лодка", "pvc boat"),
    "Квадроциклы": ("квадроцикл", "atv", "квадрик"),
    "Дорожные мотоциклы": ("дорожный мотоцикл", "street motorcycle"),
    "Внедорожные мотоциклы": (
        "эндуро",
        "кроссовый мотоцикл",
        "питбайк",
        "off-road motorcycle",
    ),
    "Снегоходы": ("снегоход", "snowmobile"),
    "Снегоуборщики": ("снегоуборщик", "snow blower", "snowblower"),
    "Мотобуксировщики": ("мотобуксировщик", "мотособака", "motorized towing"),
    "Электромотоциклы": ("электромотоцикл", "электробайк", "electric motorcycle"),
    "Гольфкары": (
        "гольфкар",
        "гольф-кар",
        "golf cart",
        "greencamel",
        "club car",
        "ezgo",
    ),
}

COLUMN_HINTS: dict[str, tuple[str, ...]] = {
    "Лодочные моторы": (
        "дейдвуд",
        "вращение винта",
        "система подъёма",
        "тип насадки",
        "топливная смесь",
    ),
    "Лодки ПВХ": (
        "плотность пвх",
        "плотность материала",
        "диаметр баллона",
        "диаметр борта",
        "тип днища",
        "пассажировместимость",
    ),
    "Квадроциклы": ("тип привода", "колесная база", "лебедка"),
    "Дорожные мотоциклы": (
        "наличие птс",
        "высота по седлу",
        "тип мотоцикла",
    ),
    "Внедорожные мотоциклы": (
        "карбюратор",
        "размер колес",
        "высота по седлу",
        "тип мотоцикла",
    ),
    "Снегоходы": ("ширина гусеницы", "длина гусеницы"),
    "Снегоуборщики": ("ширина захвата", "дальность выброса"),
    "Мотобуксировщики": ("модуль-толкач", "ширина гусеницы"),
    "Гольфкары": (
        "число мест",
        "запас хода",
        "пассажировместимость",
        "контроллер",
        "зарядное устройство",
        "педаль акселератора",
    ),
}


def _normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def detect_category(
    product_name: str,
    columns: Iterable[str],
    sheet_name: str = "",
    source_name: str = "",
) -> tuple[str, float]:
    """Return the best category and a transparent heuristic confidence."""
    name = _normalize(product_name)
    normalized_columns = " | ".join(_normalize(column) for column in columns)
    normalized_sheet = _normalize(sheet_name)
    normalized_source = _normalize(source_name)
    scores: dict[str, float] = {category: 0.0 for category in CATEGORY_RULES}

    for category, keywords in CATEGORY_RULES.items():
        normalized_category = _normalize(category)
        if normalized_source.startswith(normalized_category):
            scores[category] += 10.0
        elif normalized_category in normalized_source:
            scores[category] += 6.0
        if normalized_sheet.startswith(normalized_category):
            scores[category] += 3.0
        elif normalized_category in normalized_sheet:
            scores[category] += 1.5
        for keyword in keywords:
            if keyword in name:
                scores[category] += 2.0
        for hint in COLUMN_HINTS.get(category, ()):
            if hint in normalized_columns:
                scores[category] += 1.0

    # Common model-prefix hints are intentionally weak.
    if re.search(r"\b(honda|yamaha|suzuki|mercury|tohatsu)\s+(bf|df|f|mfs)\d+", name):
        scores["Лодочные моторы"] += 3.0

    category, score = max(scores.items(), key=lambda item: item[1])
    if score <= 0:
        return "Не определена", 0.0
    return category, min(0.98, 0.45 + score * 0.08)

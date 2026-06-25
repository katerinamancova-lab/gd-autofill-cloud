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
    "Гольфкары": ("гольфкар", "гольф-кар", "golf cart"),
}

COLUMN_HINTS: dict[str, tuple[str, ...]] = {
    "Лодочные моторы": ("мощность, л.с", "дейдвуд", "винт", "тактность"),
    "Лодки ПВХ": ("плотность пвх", "диаметр баллона", "пассажировместимость"),
    "Квадроциклы": ("тип привода", "колесная база", "лебедка"),
    "Снегоходы": ("ширина гусеницы", "длина гусеницы"),
    "Снегоуборщики": ("ширина захвата", "дальность выброса"),
    "Мотобуксировщики": ("модуль-толкач", "ширина гусеницы"),
    "Гольфкары": ("число мест", "запас хода"),
}


def _normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def detect_category(product_name: str, columns: Iterable[str]) -> tuple[str, float]:
    """Return the best category and a transparent heuristic confidence."""
    name = _normalize(product_name)
    normalized_columns = " | ".join(_normalize(column) for column in columns)
    scores: dict[str, float] = {category: 0.0 for category in CATEGORY_RULES}

    for category, keywords in CATEGORY_RULES.items():
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

import re

CATEGORY_KEYS = {
    "лодочный мотор": ["дейдвуд", "вращение винта", "тип насадки", "передачи"],
    "лодка пвх": ["тип днища", "плотность материала", "диаметр борта", "внутренняя длина"],
    "квадроцикл": ["наличие псм", "лебедка", "тип привода", "защита рук", "фаркоп"],
    "гольфкар": ["пассажировместимость", "мощность (вт)", "запас хода", "педаль акселератора"],
    "мотоцикл": ["тип мотоцикла", "наличие птс", "колеса передние", "колеса задние"],
}

BRAND_DEFAULTS = {
    "honda": ("Honda", "Япония", "Япония"),
    "tohatsu": ("Tohatsu", "Япония", "Япония"),
    "suzuki": ("Suzuki", "Япония", "Япония"),
    "yamaha": ("Yamaha", "Япония", "Япония"),
    "mercury": ("Mercury", "США", "Китай"),
    "hidea": ("Hidea", "Китай", "Китай"),
    "hdx": ("HDX", "Китай", "Китай"),
    "parsun": ("Parsun", "Китай", "Китай"),
    "sea-pro": ("Sea-Pro", "Китай", "Китай"),
    "bajaj": ("Bajaj", "Индия", "Индия"),
    "voge": ("VOGE", "Китай", "Китай"),
    "qjmotor": ("QJMotor", "Китай", "Китай"),
    "royal enfield": ("Royal Enfield", "Индия", "Индия"),
    "ktm": ("KTM", "Австрия", "Индия"),
    "benelli": ("Benelli", "Италия", "Китай"),
    "stels": ("Stels", "Россия", "Китай"),
    "linhai": ("Linhai Yamaha", "Китай", "Китай"),
    "greencamel": ("GreenCamel", "Россия", "Китай"),
    "sprmotors": ("SPRMOTORS", "Китай", "Китай"),
}


def detect_category(headers, first_names):
    joined = " ".join(str(h or "").lower() for h in headers)
    for cat, keys in CATEGORY_KEYS.items():
        if any(k in joined for k in keys):
            return cat

    names = " ".join(first_names).lower()
    if any(x in names for x in ["bf", "mfs", "mercury", "tohatsu", "hidea", "parsun", "sea-pro", "suzuki df"]):
        return "лодочный мотор"
    if any(x in names for x in ["пвх", "лодка", "нднд", "airdeck", "байкал", "сапфир"]):
        return "лодка пвх"
    if any(x in names for x in ["atv", "outlander", "segway", "linhai", "квадроцикл", "sprmotors", "blade"]):
        return "квадроцикл"
    if any(x in names for x in ["greencamel", "гольфкар", "сонора"]):
        return "гольфкар"
    return "мотоцикл"


def base_rules(product_name, category):
    spec = {}
    p = product_name.lower()

    for key, (brand, country, made) in BRAND_DEFAULTS.items():
        if key in p:
            spec["Бренд [BRAND]"] = brand
            spec["Страна бренда [BRAND_COUNTRY]"] = country
            spec["Страна производства [MANUFACTURER]"] = made
            break

    if category == "лодка пвх":
        m = re.search(r"(\d{3})", p)
        if m:
            spec["Длина, см [LENGTH_CM]"] = int(m.group(1))
        spec.setdefault("Тип лодки [TYPE_BOAT]", "Под мотор")

    spec.setdefault("Гарантия [WARRANTY]", "1 год")
    spec.setdefault("Гарантия на товар", '<span style="font-family: Acrom; font-size: 13pt;">Гарантия на товар составляет 1 год</span>')
    spec.setdefault("Скидка", 11)
    spec.setdefault("Доступное количество", 1000)
    spec.setdefault("Сортировка", 500)

    return spec

"""
Система адаптивных тем для презентаций.
Автоматический подбор цветовой палитры, шрифтов и стиля на основе тематики.
"""

from pptx.dml.color import RGBColor


# =============================================================================
# ОПРЕДЕЛЕНИЕ ТЕМ
# =============================================================================

THEMES = {
    "tech": {
        "name": "Технологии / IT / Программирование",
        "keywords": [
            "программирование", "it", "информатика", "технологии", "компьютер",
            "интернет", "сеть", "алгоритм", "данные", "искусственный интеллект",
            "машинное обучение", "кибербезопасность", "робот", "цифров", "софт",
            "hardware", "software", "web", "разработ", "код", "база данных",
            "нейросет", "блокчейн", "облачн", "devops", "python", "java",
        ],
        "fonts": {
            "heading": "Montserrat",
            "body": "Roboto",
            "fallback_heading": "Calibri",
            "fallback_body": "Arial",
        },
        "colors": {
            "bg_primary": RGBColor(0x0F, 0x0C, 0x29),
            "bg_secondary": RGBColor(0x1A, 0x1A, 0x2E),
            "accent": RGBColor(0x00, 0xD4, 0xFF),
            "accent_light": RGBColor(0x7F, 0xF0, 0xFF),
            "text_primary": RGBColor(0xFF, 0xFF, 0xFF),
            "text_secondary": RGBColor(0xE0, 0xE0, 0xE0),
            "text_muted": RGBColor(0x8A, 0x8A, 0x9A),
            "gradient_start": RGBColor(0x0F, 0x0C, 0x29),
            "gradient_end": RGBColor(0x30, 0x2B, 0x63),
        },
        "style": "dark",
    },

    "humanities": {
        "name": "Гуманитарные науки / История / Философия / Литература",
        "keywords": [
            "история", "философия", "литература", "культура", "язык",
            "лингвистик", "социолог", "психолог", "педагогик", "религи",
            "этик", "эстетик", "антрополог", "археолог", "мифолог",
            "политолог", "цивилизац", "эпоха", "общество", "традиц",
            "искусствовед", "филолог", "поэзия", "проза", "древн",
        ],
        "fonts": {
            "heading": "Georgia",
            "body": "Palatino Linotype",
            "fallback_heading": "Times New Roman",
            "fallback_body": "Times New Roman",
        },
        "colors": {
            "bg_primary": RGBColor(0xFA, 0xF3, 0xE8),
            "bg_secondary": RGBColor(0xF5, 0xEC, 0xDC),
            "accent": RGBColor(0x8B, 0x45, 0x13),
            "accent_light": RGBColor(0xC4, 0x7E, 0x4A),
            "text_primary": RGBColor(0x2C, 0x1A, 0x0E),
            "text_secondary": RGBColor(0x4A, 0x36, 0x28),
            "text_muted": RGBColor(0x7A, 0x6A, 0x5A),
            "gradient_start": RGBColor(0xFA, 0xF3, 0xE8),
            "gradient_end": RGBColor(0xF0, 0xE4, 0xD0),
        },
        "style": "light",
    },

    "science": {
        "name": "Естественные науки / Физика / Химия / Биология",
        "keywords": [
            "физика", "химия", "биология", "математика", "экология",
            "генетик", "молекул", "атом", "клетка", "эволюц",
            "эксперимент", "формула", "теория", "гипотеза", "закон",
            "энергия", "вещество", "реакция", "организм", "природ",
            "астроном", "геолог", "географ", "климат", "медицин",
        ],
        "fonts": {
            "heading": "Montserrat",
            "body": "Open Sans",
            "fallback_heading": "Calibri",
            "fallback_body": "Arial",
        },
        "colors": {
            "bg_primary": RGBColor(0x0A, 0x1A, 0x2A),
            "bg_secondary": RGBColor(0x12, 0x26, 0x3A),
            "accent": RGBColor(0x00, 0xB8, 0x7A),
            "accent_light": RGBColor(0x4E, 0xE8, 0xB0),
            "text_primary": RGBColor(0xFF, 0xFF, 0xFF),
            "text_secondary": RGBColor(0xD8, 0xE8, 0xE0),
            "text_muted": RGBColor(0x7A, 0x9A, 0x8A),
            "gradient_start": RGBColor(0x0A, 0x1A, 0x2A),
            "gradient_end": RGBColor(0x1A, 0x3A, 0x3A),
        },
        "style": "dark",
    },

    "art": {
        "name": "Искусство / Дизайн / Творчество",
        "keywords": [
            "искусство", "дизайн", "живопись", "архитектура", "скульптур",
            "фотограф", "кино", "театр", "музыка", "танец",
            "творчеств", "стиль", "мода", "композиц", "цвет",
            "графика", "анимац", "интерьер", "декор", "ремесл",
            "текстиль", "костюм", "couture", "бренд", "визуальн",
        ],
        "fonts": {
            "heading": "Raleway",
            "body": "Lato",
            "fallback_heading": "Calibri Light",
            "fallback_body": "Arial",
        },
        "colors": {
            "bg_primary": RGBColor(0x1A, 0x0A, 0x2E),
            "bg_secondary": RGBColor(0x2D, 0x13, 0x4A),
            "accent": RGBColor(0xE9, 0x4D, 0x8A),
            "accent_light": RGBColor(0xFF, 0x8C, 0xB8),
            "text_primary": RGBColor(0xFF, 0xFF, 0xFF),
            "text_secondary": RGBColor(0xF0, 0xE0, 0xF0),
            "text_muted": RGBColor(0xA0, 0x80, 0xA0),
            "gradient_start": RGBColor(0x1A, 0x0A, 0x2E),
            "gradient_end": RGBColor(0x3D, 0x1A, 0x5C),
        },
        "style": "dark",
    },

    "business": {
        "name": "Экономика / Бизнес / Менеджмент",
        "keywords": [
            "экономик", "бизнес", "менеджмент", "маркетинг", "финанс",
            "бухгалтер", "управлен", "предприним", "инвестиц", "банк",
            "рынок", "торговл", "логистик", "стратегия", "компания",
            "стартап", "прибыль", "бюджет", "налог", "аудит",
            "право", "юриспруденц", "закон", "контракт", "договор",
        ],
        "fonts": {
            "heading": "Montserrat",
            "body": "Roboto",
            "fallback_heading": "Calibri",
            "fallback_body": "Arial",
        },
        "colors": {
            "bg_primary": RGBColor(0xF8, 0xF9, 0xFA),
            "bg_secondary": RGBColor(0xEE, 0xF1, 0xF5),
            "accent": RGBColor(0x1A, 0x56, 0xDB),
            "accent_light": RGBColor(0x5B, 0x8D, 0xEF),
            "text_primary": RGBColor(0x1A, 0x1A, 0x2E),
            "text_secondary": RGBColor(0x3A, 0x3A, 0x4E),
            "text_muted": RGBColor(0x6C, 0x75, 0x7D),
            "gradient_start": RGBColor(0xF8, 0xF9, 0xFA),
            "gradient_end": RGBColor(0xE8, 0xEC, 0xF4),
        },
        "style": "light",
    },

    "minimalist": {
        "name": "Минималистичная / Универсальная",
        "keywords": [],  # Используется как fallback
        "fonts": {
            "heading": "Montserrat",
            "body": "Open Sans",
            "fallback_heading": "Calibri",
            "fallback_body": "Arial",
        },
        "colors": {
            "bg_primary": RGBColor(0xFF, 0xFF, 0xFF),
            "bg_secondary": RGBColor(0xF5, 0xF5, 0xF5),
            "accent": RGBColor(0x2D, 0x2D, 0x2D),
            "accent_light": RGBColor(0x5A, 0x5A, 0x5A),
            "text_primary": RGBColor(0x1A, 0x1A, 0x1A),
            "text_secondary": RGBColor(0x3A, 0x3A, 0x3A),
            "text_muted": RGBColor(0x7A, 0x7A, 0x7A),
            "gradient_start": RGBColor(0xFF, 0xFF, 0xFF),
            "gradient_end": RGBColor(0xF0, 0xF0, 0xF0),
        },
        "style": "light",
    },
}


# =============================================================================
# ОПРЕДЕЛЕНИЕ ТЕМЫ ПО КЛЮЧЕВЫМ СЛОВАМ
# =============================================================================

def detect_theme(title: str, description: str = "") -> str:
    """
    Определяет подходящую тему на основе заголовка и описания.
    Возвращает ключ темы из словаря THEMES.
    """
    text = (title + " " + description).lower()

    scores = {}
    for theme_key, theme_data in THEMES.items():
        if theme_key == "minimalist":
            continue
        score = 0
        for keyword in theme_data["keywords"]:
            if keyword.lower() in text:
                score += 1
        if score > 0:
            scores[theme_key] = score

    if not scores:
        return "minimalist"

    return max(scores, key=scores.get)


def get_theme(theme_key: str) -> dict:
    """Получает данные темы по ключу."""
    return THEMES.get(theme_key, THEMES["minimalist"])


def list_themes() -> list:
    """Возвращает список доступных тем."""
    return [(key, data["name"]) for key, data in THEMES.items()]


# =============================================================================
# ИНФОРМАЦИЯ О ВУЗЕ
# =============================================================================

UNIVERSITY_INFO = {
    "full_name": "Российский государственный университет имени А.Н. Косыгина\n(Технологии. Дизайн. Искусство)",
    "short_name": "РГУ им. А.Н. Косыгина",
    "branch": "Тверской филиал",
    "branch_full": "Тверской филиал федерального государственного бюджетного\nобразовательного учреждения высшего образования\n«Российский государственный университет имени А.Н. Косыгина\n(Технологии. Дизайн. Искусство)»",
    "ministry": "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ\nРОССИЙСКОЙ ФЕДЕРАЦИИ",
    "city": "Тверь",
    "year": 2026,
}


# =============================================================================
# СТАНДАРТЫ ОФОРМЛЕНИЯ (ГОСТ 7.32-2017)
# =============================================================================

GOST_STANDARDS = {
    "font_main": "Times New Roman",
    "font_size_main": 14,
    "font_size_heading1": 16,
    "font_size_heading2": 14,
    "line_spacing": 1.5,
    "margin_left": 30,      # мм
    "margin_right": 15,     # мм
    "margin_top": 20,       # мм
    "margin_bottom": 20,    # мм
    "first_line_indent": 12.5,  # мм (1.25 см)
}

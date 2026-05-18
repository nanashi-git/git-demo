#!/usr/bin/env python3
"""
Презентация: Антонис Ван Дейк — избранные картины
Предмет: История искусств
Современный дизайн с картинками
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from datetime import datetime
from PIL import Image as PILImage
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = os.path.join(SCRIPT_DIR, "images")

# Палитра: тёмная с золотыми акцентами (стиль барокко)
C = {
    "bg": RGBColor(0x12, 0x0B, 0x1E),
    "bg2": RGBColor(0x1C, 0x12, 0x30),
    "card": RGBColor(0x1E, 0x14, 0x35),
    "accent": RGBColor(0xC9, 0x9A, 0x2E),
    "accent_lt": RGBColor(0xE8, 0xC5, 0x6D),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "light": RGBColor(0xE8, 0xE0, 0xD0),
    "muted": RGBColor(0x8A, 0x7A, 0x6A),
    "grad1": RGBColor(0x08, 0x05, 0x12),
    "grad2": RGBColor(0x22, 0x15, 0x3A),
}

F_HEAD = "Georgia"
F_BODY = "Calibri"


def gradient_bg(slide, c1, c2):
    fill = slide.background.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = c1
    fill.gradient_stops[1].color.rgb = c2


def solid_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def oval(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def text(slide, l, t, w, h, txt, font=F_BODY, sz=Pt(14), color=None,
         bold=False, italic=False, align=PP_ALIGN.LEFT):
    if color is None:
        color = C["light"]
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.name = font
    p.font.size = sz
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.alignment = align
    return tf


def multitext(slide, l, t, w, h, lines, font=F_BODY, sz=Pt(13), color=None, spacing=Pt(6)):
    if color is None:
        color = C["light"]
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = font
        p.font.size = sz
        p.font.color.rgb = color
        p.space_before = spacing
    return tf


def get_image_size(img_name):
    """Получает размеры изображения"""
    path = os.path.join(IMAGES_DIR, img_name)
    if not os.path.exists(path):
        return (800, 1000)  # fallback
    with PILImage.open(path) as img:
        return img.size  # (width, height)


def add_image_fit(slide, img_name, left, top, max_width, max_height):
    """
    Добавляет картину, вписывая в область max_width x max_height
    с сохранением пропорций. Центрирует по горизонтали и вертикали.
    """
    path = os.path.join(IMAGES_DIR, img_name)
    if not os.path.exists(path):
        print(f"  WARNING: Image not found: {path}")
        return None

    img_w, img_h = get_image_size(img_name)
    aspect = img_w / img_h

    # Вычисляем размер, вписывая в область
    max_w_inches = max_width / 914400  # Emu to inches
    max_h_inches = max_height / 914400

    if aspect > (max_w_inches / max_h_inches):
        # Картина шире — ограничиваем по ширине
        final_w = max_width
        final_h = int(max_width / aspect)
    else:
        # Картина выше — ограничиваем по высоте
        final_h = max_height
        final_w = int(max_height * aspect)

    # Центрирование
    offset_x = (max_width - final_w) // 2
    offset_y = (max_height - final_h) // 2

    pic = slide.shapes.add_picture(path, left + offset_x, top + offset_y,
                                   width=final_w, height=final_h)
    return pic


def slide_num(slide, n):
    text(slide, Inches(9.3), Inches(7.05), Inches(0.5), Inches(0.3),
         str(n), sz=Pt(9), color=C["muted"], align=PP_ALIGN.RIGHT)


def painting_slide(prs, num, title, year, img_file, facts):
    """Создаёт слайд с картиной: изображение слева, текст справа"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    solid_bg(slide, C["bg"])

    # Тонкая золотая линия сверху
    rect(slide, Inches(0), Inches(0), Inches(10), Pt(3), C["accent"])

    # Подложка для картины
    card_left = Inches(0.3)
    card_top = Inches(0.5)
    card_w = Inches(4.0)
    card_h = Inches(6.6)
    rect(slide, card_left, card_top, card_w, card_h, C["card"])

    # Картина — вписываем с отступами внутри подложки
    img_padding = Inches(0.15)
    add_image_fit(slide, img_file,
                  left=card_left + img_padding,
                  top=card_top + img_padding,
                  max_width=Emu(card_w - img_padding * 2),
                  max_height=Emu(card_h - img_padding * 2))

    # Текстовая часть справа (начинается после картины)
    text_left = Inches(4.6)
    text_width = Inches(5.1)

    # Заголовок
    text(slide, text_left, Inches(0.6), text_width, Inches(0.8),
         title, font=F_HEAD, sz=Pt(20), color=C["white"], bold=True)

    # Год / техника
    text(slide, text_left, Inches(1.35), text_width, Inches(0.4),
         year, sz=Pt(12), color=C["accent_lt"], italic=True)

    # Акцентная линия
    rect(slide, text_left, Inches(1.8), Inches(1.5), Pt(2), C["accent"])

    # Описание
    multitext(slide, text_left, Inches(2.0), text_width, Inches(5.0),
              facts, sz=Pt(12), spacing=Pt(8))

    # Номер слайда
    slide_num(slide, num)

    return slide


def main():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # ============================
    # СЛАЙД 1: ТИТУЛЬНЫЙ
    # ============================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    gradient_bg(slide, C["grad1"], C["grad2"])

    # Декоративные элементы
    oval(slide, Inches(8.0), Inches(-1.5), Inches(3.5), Inches(3.5), C["accent"])
    oval(slide, Inches(-1.0), Inches(6.0), Inches(2.0), Inches(2.0), C["accent_lt"])

    # Автопортрет справа (вписан в зону)
    add_image_fit(slide, "self_portrait.jpg",
                  left=Emu(Inches(6.8)), top=Emu(Inches(1.2)),
                  max_width=Emu(Inches(2.8)), max_height=Emu(Inches(5.5)))

    # Вуз
    text(slide, Inches(0.8), Inches(0.5), Inches(5.5), Inches(0.4),
         "Тверской филиал РГУ им. А.Н. Косыгина", sz=Pt(10), color=C["muted"])

    # Предмет
    text(slide, Inches(0.8), Inches(2.0), Inches(5), Inches(0.4),
         "ИСТОРИЯ ИСКУССТВ", sz=Pt(12), color=C["accent_lt"], italic=True)

    # Золотая линия
    rect(slide, Inches(0.8), Inches(2.55), Inches(3.0), Pt(4), C["accent"])

    # Заголовок
    text(slide, Inches(0.8), Inches(2.8), Inches(5.5), Inches(1.8),
         "АНТОНИС\nВАН ДЕЙК", font=F_HEAD, sz=Pt(42), color=C["white"], bold=True)

    # Подзаголовок
    text(slide, Inches(0.8), Inches(4.6), Inches(5.5), Inches(0.6),
         "Избранные произведения", sz=Pt(16), color=C["light"])

    # Даты жизни
    text(slide, Inches(0.8), Inches(5.3), Inches(5), Inches(0.4),
         "1599 – 1641", sz=Pt(14), color=C["accent_lt"])

    # Автор
    text(slide, Inches(0.8), Inches(6.5), Inches(5), Inches(0.4),
         f"Студент 1 курса  •  {datetime.now().year}", sz=Pt(10), color=C["muted"])

    # ============================
    # СЛАЙД 2: БИОГРАФИЯ
    # ============================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    solid_bg(slide, C["bg"])

    rect(slide, Inches(0), Inches(0), Inches(10), Pt(3), C["accent"])

    # Автопортрет слева — подложка
    card_w = Inches(3.0)
    card_h = Inches(6.4)
    rect(slide, Inches(0.3), Inches(0.7), card_w, card_h, C["card"])

    # Вписываем автопортрет
    add_image_fit(slide, "self_portrait.jpg",
                  left=Emu(Inches(0.4)), top=Emu(Inches(0.8)),
                  max_width=Emu(Inches(2.8)), max_height=Emu(Inches(6.2)))

    # Текст справа
    text(slide, Inches(3.6), Inches(0.5), Inches(6.0), Inches(0.7),
         "Антонис Ван Дейк", font=F_HEAD, sz=Pt(26), color=C["white"], bold=True)

    text(slide, Inches(3.6), Inches(1.15), Inches(6.0), Inches(0.4),
         "22 марта 1599, Антверпен — 9 декабря 1641, Лондон",
         sz=Pt(11), color=C["accent_lt"], italic=True)

    rect(slide, Inches(3.6), Inches(1.6), Inches(1.5), Pt(2), C["accent"])

    bio_lines = [
        "Фламандский живописец, один из величайших",
        "портретистов эпохи барокко",
        "",
        "• Ученик и помощник Питера Пауля Рубенса",
        "• Итальянский период (1621–1627): Генуя, Рим",
        "• Придворный художник Карла I с 1632 года",
        "• Создал более 900 картин за 42 года жизни",
        "• Революционизировал жанр парадного портрета",
        "• Определил развитие английской портретной",
        "  живописи на 150 лет вперёд",
    ]
    multitext(slide, Inches(3.6), Inches(1.8), Inches(6.0), Inches(5.4),
              bio_lines, sz=Pt(13), spacing=Pt(9))

    slide_num(slide, 2)

    # ============================
    # СЛАЙДЫ 3–10: КАРТИНЫ
    # ============================

    painting_slide(prs, 3,
        "Семейный портрет",
        "Ок. 1620–1621 • Холст, масло • 113,5 × 93,5 см",
        "family_portrait.jpg",
        [
            "Государственный Эрмитаж, Санкт-Петербург",
            "",
            "Одна из самых проникновенных работ раннего",
            "периода. Изображена супружеская пара с",
            "дочерью на коленях у матери.",
            "",
            "Компактная композиция создаёт атмосферу",
            "доверительной близости и семейного согласия.",
            "",
            "Ван Дейку было всего 21 год — но уже видно",
            "влияние Рубенса в колористике и свободной,",
            "уверенной манере письма.",
        ])

    painting_slide(prs, 4,
        "Святой Мартин и нищий",
        "Ок. 1618 • Дерево, масло • 172 × 158 см",
        "st_martin.jpg",
        [
            "Церковь Св. Мартина, Завентем (Бельгия)",
            "",
            "Алтарный образ, написанный 19-летним Ван",
            "Дейком. Сюжет: римский воин Мартин Турский",
            "разрезает свой плащ, чтобы поделиться",
            "с замерзающим нищим.",
            "",
            "Заказ предназначался Рубенсу, но был",
            "передан талантливому ученику.",
            "",
            "Мощная барочная диагональ и драматизм",
            "свидетельствуют о раннем мастерстве.",
        ])

    painting_slide(prs, 5,
        "Портрет кардинала Бентивольо",
        "Ок. 1623 • Холст, масло",
        "bentivoglio.jpg",
        [
            "Галерея Палатина, Палаццо Питти, Флоренция",
            "",
            "Написан в итальянский период. Кардинал",
            "Гвидо Бентивольо — дипломат, связанный",
            "с Фландрией.",
            "",
            "«Весь Рим устремился смотреть это чудо",
            "искусства, и каждый хотел быть написанным",
            "рукой нашего художника»",
            "",
            "Работа, вдохновлённая Тицианом, утвердила",
            "Ван Дейка как ведущего портретиста эпохи.",
        ])

    painting_slide(prs, 6,
        "Автопортрет",
        "Ок. 1620–1621 • Холст, масло • 119,7 × 87,9 см",
        "self_portrait.jpg",
        [
            "Метрополитен-музей, Нью-Йорк",
            "",
            "Написан зимой 1620–1621 в Лондоне.",
            "Ван Дейк изображает себя как светского",
            "джентльмена — без палитры и кистей.",
            "",
            "Небрежная поза руки у подбородка",
            "подчёркивает аристократизм.",
            "",
            "Виртуозная кисть передаёт блеск шёлка",
            "и сияние молодой кожи 21-летнего мастера.",
        ])

    painting_slide(prs, 7,
        "Портрет Марии Луизы де Тассис",
        "Ок. 1629–1630 • Холст, масло",
        "maria_de_tassis.jpg",
        [
            "Коллекция князей Лихтенштейн, Вадуц",
            "",
            "Мария Луиза (1611–1638) — дочь Антонио",
            "де Тассис из Антверпена. Семья создала",
            "первую почтовую систему Европы.",
            "",
            "Изображена в возрасте около 19 лет.",
            "Роскошное платье, кружева, жемчуг —",
            "утончённая красота аристократки.",
            "",
            "Эталон мастерства Ван Дейка в передаче",
            "тканей, украшений и женственности.",
        ])

    painting_slide(prs, 8,
        "Портрет сэра Томаса Чалонера",
        "Конец 1630-х • Холст, масло",
        "chaloner.jpg",
        [
            "Государственный Эрмитаж, Санкт-Петербург",
            "",
            "Одна из последних работ мастера.",
            "Сэр Томас Чалонер — придворный Карла I.",
            "",
            "Художник честно передаёт стареющее лицо:",
            "дряблая кожа, румянец — без лести,",
            "глубокий психологизм.",
            "",
            "Считается одним из лучших полотен",
            "позднего периода. Свободная, уверенная",
            "манера письма зрелого мастера.",
        ])

    painting_slide(prs, 9,
        "Портрет Карла I на охоте",
        "Ок. 1635 • Холст, масло • 266 × 207 см",
        "charles_hunt.jpg",
        [
            "Музей Лувр, Париж",
            "",
            "Шедевр и жемчужина Лувра с 1793 года.",
            "Карл I в гражданской одежде у лошади.",
            "",
            "«Тонкий компромисс между джентльменской",
            "небрежностью и королевской уверенностью»",
            "",
            "Король был невысок (163 см), но Ван Дейк",
            "с помощью композиции создаёт впечатление",
            "величественной фигуры.",
        ])

    painting_slide(prs, 10,
        "Портрет Дигби и Рассела",
        "Ок. 1637 • Холст, масло",
        "digby_russell.jpg",
        [
            "Коллекция Спенсеров, Олторп, Англия",
            "",
            "Двойной парадный портрет молодых",
            "английских аристократов:",
            "",
            "• Джордж Дигби (1612–1677) — 2-й граф",
            "  Бристоль, политик, поэт, драматург",
            "• Уильям Рассел (1616–1700) — будущий",
            "  1-й герцог Бедфорд",
            "",
            "Элегантная непринуждённость поз —",
            "типичный «friendship portrait».",
        ])

    # ============================
    # СЛАЙД 11: ФИНАЛЬНЫЙ
    # ============================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    gradient_bg(slide, C["grad1"], C["grad2"])

    # Декор
    oval(slide, Inches(7.5), Inches(4.5), Inches(3.5), Inches(3.5), C["accent"])
    oval(slide, Inches(-0.5), Inches(-0.5), Inches(1.5), Inches(1.5), C["accent_lt"])

    # Текст вверху
    text(slide, Inches(1.0), Inches(1.5), Inches(8), Inches(1.2),
         "СПАСИБО ЗА ВНИМАНИЕ", font=F_HEAD, sz=Pt(36),
         color=C["white"], bold=True, align=PP_ALIGN.LEFT)

    rect(slide, Inches(1.0), Inches(2.8), Inches(2.5), Pt(3), C["accent"])

    text(slide, Inches(1.0), Inches(3.1), Inches(6), Inches(0.5),
         "Готов ответить на вопросы", sz=Pt(18), color=C["accent_lt"])

    text(slide, Inches(1.0), Inches(3.8), Inches(6), Inches(0.4),
         "Тверской филиал РГУ им. А.Н. Косыгина  •  История искусств",
         sz=Pt(11), color=C["muted"])

    # Мини-галерея внизу — 4 картины с правильными пропорциями
    gallery_imgs = ["family_portrait.jpg", "bentivoglio.jpg",
                    "charles_hunt.jpg", "maria_de_tassis.jpg"]
    gallery_h = Inches(2.5)
    x_pos = Inches(0.5)
    for img in gallery_imgs:
        img_w, img_h = get_image_size(img)
        aspect = img_w / img_h
        pic_h = gallery_h
        pic_w = int(Emu(gallery_h) * aspect)
        path = os.path.join(IMAGES_DIR, img)
        if os.path.exists(path):
            slide.shapes.add_picture(path, x_pos, Inches(4.6), height=gallery_h)
            # Сдвигаем x на ширину картины + отступ
            actual_w = gallery_h * aspect / 914400  # в дюймах
            x_pos += Inches(actual_w + 0.15)

    # ============================
    # СОХРАНЕНИЕ
    # ============================
    output = os.path.join(SCRIPT_DIR, "Van_Dyck_Presentation.pptx")
    prs.save(output)
    print(f"\n  Презентация создана: {output}")
    print(f"  Слайдов: 11 (титульный + биография + 8 картин + финальный)")
    print(f"  Дизайн: тёмный с золотыми акцентами, изображения картин")
    print(f"  Шрифты: Georgia + Calibri")


if __name__ == "__main__":
    main()

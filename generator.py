#!/usr/bin/env python3
"""
ГЕНЕРАТОР ПРЕЗЕНТАЦИЙ И РЕФЕРАТОВ
==================================
Универсальный скрипт для быстрого создания:
- Стильных презентаций (.pptx)
- Рефератов/докладов (.docx)

Использование:
    python generator.py presentation "Тема презентации" --author "Имя" --slides 10
    python generator.py referat "Тема реферата" --author "Имя" --group "XX-XXX"
"""

import argparse
import os
import sys
from datetime import datetime


def create_presentation(title, author="Студент", group="XX-XXX", slides_count=6, subtitle="", output=None):
    """Создаёт стильную презентацию"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    COLORS = {
        'primary': RGBColor(0x1A, 0x1A, 0x2E),
        'secondary': RGBColor(0x16, 0x21, 0x3E),
        'accent': RGBColor(0x0F, 0x8B, 0x8D),
        'accent_light': RGBColor(0x53, 0xCE, 0xD1),
        'white': RGBColor(0xFF, 0xFF, 0xFF),
        'light_gray': RGBColor(0xE8, 0xE8, 0xE8),
        'dark_gray': RGBColor(0x6C, 0x75, 0x7D),
        'gradient_start': RGBColor(0x0F, 0x0C, 0x29),
        'gradient_end': RGBColor(0x30, 0x2B, 0x63),
    }

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    def add_gradient_bg(slide, c1, c2):
        bg = slide.background
        fill = bg.fill
        fill.gradient()
        fill.gradient_stops[0].color.rgb = c1
        fill.gradient_stops[1].color.rgb = c2

    def add_solid_bg(slide, color):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_line(slide, left, top, width, height, color):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.fill.background()

    def add_circle(slide, left, top, size, color):
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.fill.background()

    # === ТИТУЛЬНЫЙ СЛАЙД ===
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, COLORS['gradient_start'], COLORS['gradient_end'])
    add_circle(slide, Inches(8.5), Inches(-1), Inches(3), COLORS['accent'])
    add_circle(slide, Inches(-0.5), Inches(5.5), Inches(2), COLORS['accent_light'])
    add_line(slide, Inches(1.5), Inches(2.8), Inches(2), Pt(4), COLORS['accent'])

    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(3.0), Inches(7), Inches(1.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title.upper()
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['white']

    if subtitle:
        txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.5), Inches(7), Inches(1))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = subtitle
        p2.font.size = Pt(18)
        p2.font.color.rgb = COLORS['light_gray']

    txBox3 = slide.shapes.add_textbox(Inches(1.5), Inches(6.2), Inches(7), Inches(0.8))
    tf3 = txBox3.text_frame
    p3 = tf3.paragraphs[0]
    p3.text = f"Автор: {author}  |  Группа: {group}  |  {datetime.now().year}"
    p3.font.size = Pt(12)
    p3.font.color.rgb = COLORS['dark_gray']

    # === СЛАЙДЫ КОНТЕНТА ===
    for i in range(1, slides_count - 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_solid_bg(slide, COLORS['primary'])

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.4), Inches(0.8))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = f"Слайд {i + 1}: Заголовок"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']

        add_line(slide, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3), COLORS['accent'])

        txBox2 = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(8.4), Inches(5.5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = "[Содержимое слайда]"
        p2.font.size = Pt(16)
        p2.font.color.rgb = COLORS['light_gray']

    # === ФИНАЛЬНЫЙ СЛАЙД ===
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, COLORS['gradient_start'], COLORS['gradient_end'])
    add_circle(slide, Inches(7.5), Inches(4.5), Inches(3.5), COLORS['accent'])

    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(7), Inches(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "СПАСИБО ЗА ВНИМАНИЕ"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLORS['white']

    txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.5), Inches(6), Inches(1))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = "Вопросы?"
    p2.font.size = Pt(20)
    p2.font.color.rgb = COLORS['accent_light']

    # Сохранение
    if not output:
        safe_title = "".join(c if c.isalnum() or c in ' _-' else '' for c in title)[:50]
        output = f"{safe_title}.pptx"
    prs.save(output)
    print(f"Презентация создана: {output}")
    print(f"  Тема: {title}")
    print(f"  Слайдов: {slides_count}")
    print(f"  Автор: {author}")
    return output



def create_referat(title, author="Студент", group="XX-XXX", university="НАЗВАНИЕ УНИВЕРСИТЕТА",
                   department="Название кафедры", discipline="Название дисциплины",
                   city="Москва", output=None):
    """Создаёт реферат с оформлением по ГОСТ"""
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # Настройка полей
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)

    # Настройка стиля Normal
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.first_line_indent = Cm(1.25)

    # Heading 1
    h1 = doc.styles['Heading 1']
    h1.font.name = 'Times New Roman'
    h1.font.size = Pt(16)
    h1.font.bold = True
    from docx.shared import RGBColor
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(24)
    h1.paragraph_format.space_after = Pt(12)
    h1.paragraph_format.first_line_indent = Cm(0)

    # Heading 2
    h2 = doc.styles['Heading 2']
    h2.font.name = 'Times New Roman'
    h2.font.size = Pt(14)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 0, 0)
    h2.paragraph_format.space_before = Pt(18)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.first_line_indent = Cm(1.25)

    year = datetime.now().year

    # === ТИТУЛЬНЫЙ ЛИСТ ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run('МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ\nРОССИЙСКОЙ ФЕДЕРАЦИИ')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.first_line_indent = Cm(0)
    p2.paragraph_format.space_before = Pt(12)
    run2 = p2.add_run(f'«{university}»')
    run2.font.size = Pt(14)
    run2.font.bold = True
    run2.font.name = 'Times New Roman'

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.first_line_indent = Cm(0)
    p3.paragraph_format.space_before = Pt(24)
    run3 = p3.add_run(f'Кафедра «{department}»')
    run3.font.size = Pt(14)
    run3.font.name = 'Times New Roman'

    for _ in range(4):
        doc.add_paragraph().paragraph_format.first_line_indent = Cm(0)

    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.first_line_indent = Cm(0)
    run5 = p5.add_run('РЕФЕРАТ')
    run5.font.size = Pt(18)
    run5.font.bold = True
    run5.font.name = 'Times New Roman'

    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.first_line_indent = Cm(0)
    p6.paragraph_format.space_before = Pt(12)
    run6 = p6.add_run(f'по дисциплине «{discipline}»')
    run6.font.size = Pt(14)
    run6.font.name = 'Times New Roman'

    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p7.paragraph_format.first_line_indent = Cm(0)
    p7.paragraph_format.space_before = Pt(12)
    run7 = p7.add_run(f'на тему: «{title}»')
    run7.font.size = Pt(14)
    run7.font.bold = True
    run7.font.name = 'Times New Roman'

    for _ in range(5):
        doc.add_paragraph().paragraph_format.first_line_indent = Cm(0)

    p8 = doc.add_paragraph()
    p8.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p8.paragraph_format.first_line_indent = Cm(0)
    run8 = p8.add_run(f'Выполнил: студент группы {group}\n{author}')
    run8.font.size = Pt(14)
    run8.font.name = 'Times New Roman'

    for _ in range(4):
        doc.add_paragraph().paragraph_format.first_line_indent = Cm(0)

    p10 = doc.add_paragraph()
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10.paragraph_format.first_line_indent = Cm(0)
    run10 = p10.add_run(f'{city} — {year}')
    run10.font.size = Pt(14)
    run10.font.name = 'Times New Roman'

    doc.add_page_break()

    # === СОДЕРЖАНИЕ ===
    doc.add_heading('СОДЕРЖАНИЕ', level=1)
    toc_items = [
        'Введение',
        '1. Первый раздел',
        '2. Второй раздел',
        '3. Третий раздел',
        'Заключение',
        'Список использованных источников',
    ]
    for item in toc_items:
        p = doc.add_paragraph(item)
        p.paragraph_format.first_line_indent = Cm(0)

    doc.add_page_break()

    # === ВВЕДЕНИЕ ===
    doc.add_heading('ВВЕДЕНИЕ', level=1)
    doc.add_paragraph(
        f'Актуальность темы «{title}» обусловлена тем, что... '
        '[Описать актуальность темы. Объём: 1-2 страницы.]'
    )
    p_goal = doc.add_paragraph()
    run_g = p_goal.add_run('Цель работы')
    run_g.bold = True
    p_goal.add_run(' — [сформулировать цель].')

    p_tasks = doc.add_paragraph()
    run_t = p_tasks.add_run('Задачи:')
    run_t.bold = True

    for task in ['изучить...', 'проанализировать...', 'рассмотреть...', 'сделать выводы...']:
        doc.add_paragraph(f'— {task}')

    doc.add_page_break()

    # === ОСНОВНАЯ ЧАСТЬ ===
    doc.add_heading('1. ПЕРВЫЙ РАЗДЕЛ', level=1)
    doc.add_paragraph('[Текст первого раздела]')
    doc.add_page_break()

    doc.add_heading('2. ВТОРОЙ РАЗДЕЛ', level=1)
    doc.add_paragraph('[Текст второго раздела]')
    doc.add_page_break()

    doc.add_heading('3. ТРЕТИЙ РАЗДЕЛ', level=1)
    doc.add_paragraph('[Текст третьего раздела]')
    doc.add_page_break()

    # === ЗАКЛЮЧЕНИЕ ===
    doc.add_heading('ЗАКЛЮЧЕНИЕ', level=1)
    doc.add_paragraph('[Выводы по работе. Объём: 1-2 страницы.]')
    doc.add_page_break()

    # === СПИСОК ЛИТЕРАТУРЫ ===
    doc.add_heading('СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', level=1)
    refs = [
        'Фамилия, И.О. Название / И.О. Фамилия. — М.: Издательство, 2023. — 200 с.',
        'Фамилия, И.О. Статья // Журнал. — 2024. — № 1. — С. 10-15.',
        'Ресурс [Электронный ресурс]. — URL: https://example.com (дата обращения: 01.01.2026).',
    ]
    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph(f'{i}. {ref}')

    # Сохранение
    if not output:
        safe_title = "".join(c if c.isalnum() or c in ' _-' else '' for c in title)[:50]
        output = f"Реферат - {safe_title}.docx"
    doc.save(output)
    print(f"Реферат создан: {output}")
    print(f"  Тема: {title}")
    print(f"  Автор: {author}")
    print(f"  Группа: {group}")
    return output


def main():
    parser = argparse.ArgumentParser(
        description='Генератор презентаций и рефератов',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python generator.py presentation "Искусственный интеллект" --author "Иванов И.И." --slides 8
  python generator.py referat "Квантовые вычисления" --author "Петров П.П." --group "ИТ-301"
        """
    )
    subparsers = parser.add_subparsers(dest='command', help='Тип документа')

    # Презентация
    pres_parser = subparsers.add_parser('presentation', aliases=['pres', 'p'],
                                         help='Создать презентацию (.pptx)')
    pres_parser.add_argument('title', help='Тема презентации')
    pres_parser.add_argument('--author', default='Студент', help='Имя автора')
    pres_parser.add_argument('--group', default='XX-XXX', help='Номер группы')
    pres_parser.add_argument('--slides', type=int, default=6, help='Количество слайдов (по умолчанию 6)')
    pres_parser.add_argument('--subtitle', default='', help='Подзаголовок')
    pres_parser.add_argument('--output', '-o', help='Имя выходного файла')

    # Реферат
    ref_parser = subparsers.add_parser('referat', aliases=['ref', 'r'],
                                        help='Создать реферат (.docx)')
    ref_parser.add_argument('title', help='Тема реферата')
    ref_parser.add_argument('--author', default='Студент', help='Имя автора')
    ref_parser.add_argument('--group', default='XX-XXX', help='Номер группы')
    ref_parser.add_argument('--university', default='НАЗВАНИЕ УНИВЕРСИТЕТА', help='Название вуза')
    ref_parser.add_argument('--department', default='Название кафедры', help='Кафедра')
    ref_parser.add_argument('--discipline', default='Название дисциплины', help='Дисциплина')
    ref_parser.add_argument('--city', default='Москва', help='Город')
    ref_parser.add_argument('--output', '-o', help='Имя выходного файла')

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    if args.command in ('presentation', 'pres', 'p'):
        create_presentation(
            title=args.title,
            author=args.author,
            group=args.group,
            slides_count=args.slides,
            subtitle=args.subtitle,
            output=args.output
        )
    elif args.command in ('referat', 'ref', 'r'):
        create_referat(
            title=args.title,
            author=args.author,
            group=args.group,
            university=args.university,
            department=args.department,
            discipline=args.discipline,
            city=args.city,
            output=args.output
        )


if __name__ == "__main__":
    main()

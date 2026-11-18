# -*- coding: utf-8 -*-
"""
Builder for Document 3: Functional Specification (Функциональная спецификация)
for "NewFastPizza - Pizza Analysis System".
"""
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_document():
    doc = docx.Document()
    
    # 1. Page Setup (A4, Left 30mm, Right 15mm, Top 20mm, Bottom 20mm)
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.left_margin = Inches(1.18)   # 30 mm
        section.right_margin = Inches(0.59)  # 15 mm
        section.top_margin = Inches(0.79)    # 20 mm
        section.bottom_margin = Inches(0.79) # 20 mm
        
        # Configure footer page numbers
        footer = section.footer
        p_ft = footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_ft.paragraph_format.space_before = Pt(0)
        p_ft.paragraph_format.space_after = Pt(0)
        r_ft = p_ft.add_run()
        r_ft.font.name = "Times New Roman"
        r_ft.font.size = Pt(9.5)
        r_ft.font.color.rgb = RGBColor(100, 116, 139)
        fld_xml = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        r_ft._r.append(fld_xml)

    # Styling helper functions
    def set_cell_background(cell, color_hex):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_table_borders(table, color="CBD5E1"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/><w:left w:val="none"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/><w:right w:val="none"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/><w:insideV w:val="none"/></w:tblBorders>')
        tblPr.append(borders)

    def p_body(text, bold_prefix="", space_before=0, space_after=6, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(11)
            r_pre.font.color.rgb = RGBColor(15, 23, 42)
        r_txt = p.add_run(text)
        r_txt.italic = italic
        r_txt.font.name = "Times New Roman"
        r_txt.font.size = Pt(11)
        r_txt.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def p_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(15, 23, 42)
        return p

    def p_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12.5)
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def p_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(11.5)
        run.font.color.rgb = RGBColor(51, 65, 85)
        return p

    def p_bullet(title, desc):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r1 = p.add_run(title)
        r1.bold = True
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(11)
        if desc:
            r2 = p.add_run(" " + desc)
            r2.font.name = "Times New Roman"
            r2.font.size = Pt(11)
        return p

    def p_callout(text, title="ВАЖНОЕ ПРАВИЛО ПРОЕКТИРОВАНИЯ:", bg_color="FEF3C7", border_color="D97706"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.4)
        set_cell_background(cell, bg_color)
        set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
        tcPr.append(borders)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        if title:
            r_title = p.add_run(title + "\n")
            r_title.bold = True
            r_title.font.name = "Times New Roman"
            r_title.font.size = Pt(10.5)
            r_title.font.color.rgb = RGBColor(146, 64, 14)
        r_body = p.add_run(text)
        r_body.font.name = "Times New Roman"
        r_body.font.size = Pt(10)
        r_body.font.color.rgb = RGBColor(30, 41, 59)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def p_image(img_path, caption, width_in=6.2):
        if not os.path.exists(img_path):
            print(f"Warning: {img_path} not found!")
            return
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        p_img.paragraph_format.keep_with_next = True
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(width_in))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(0)
        p_cap.paragraph_format.space_after = Pt(12)
        run_cap = p_cap.add_run(caption)
        run_cap.italic = True
        run_cap.font.name = "Times New Roman"
        run_cap.font.size = Pt(9.5)
        run_cap.font.color.rgb = RGBColor(71, 85, 105)

    def p_table(headers, rows_data, col_widths=None):
        tbl = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_borders(tbl)
        
        hdr_row = tbl.rows[0]
        hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        for idx, h_text in enumerate(headers):
            cell = hdr_row.cells[idx]
            if col_widths and idx < len(col_widths):
                cell.width = Inches(col_widths[idx])
            set_cell_background(cell, "F1F5F9")
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(h_text)
            r.bold = True
            r.font.name = "Times New Roman"
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(15, 23, 42)

        for r_idx, r_data in enumerate(rows_data):
            row = tbl.rows[r_idx + 1]
            bg = "FFFFFF" if r_idx % 2 == 0 else "F8FAFC"
            for c_idx, val in enumerate(r_data):
                cell = row.cells[c_idx]
                if col_widths and c_idx < len(col_widths):
                    cell.width = Inches(col_widths[c_idx])
                set_cell_background(cell, bg)
                set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx > 0 else WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.1
                r = p.add_run(str(val))
                r.font.name = "Times New Roman"
                r.font.size = Pt(9.0)
                r.font.color.rgb = RGBColor(30, 41, 59)
                
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # =========================================================================
    # ТИТУЛЬНЫЙ ЛИСТ
    # =========================================================================
    p0 = doc.add_paragraph()
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(2)
    r = p0.add_run("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(10)
    r = p0.add_run("Федеральное государственное автономное образовательное учреждение высшего образования\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(10)
    r = p0.add_run("«Национальный исследовательский Нижегородский государственный университет\nим. Н.И. Лобачевского»\n")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(10.5)

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(4)
    p_inst.paragraph_format.space_after = Pt(2)
    r = p_inst.add_run("Институт информационных технологий, математики и механики\nКафедра математического обеспечения ЭВМ")
    r.font.name = "Times New Roman"; r.font.size = Pt(10.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_course.paragraph_format.space_before = Pt(0)
    p_course.paragraph_format.space_after = Pt(2)
    r = p_course.add_run("УЧЕБНЫЙ КУРС\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(12)
    r = p_course.add_run("«Объектно-ориентированный анализ и проектирование»\n")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(13)
    r = p_course.add_run("(«Проектирование и архитектура программных систем»)\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)
    r = p_course.add_run("Направление подготовки: 09.03.04 — Программная инженерия")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(28)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r = p_title.add_run("ФУНКЦИОНАЛЬНАЯ СПЕЦИФИКАЦИЯ\n")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(17); r.font.color.rgb = RGBColor(15, 23, 42)
    r = p_title.add_run("по проекту «Система заказа пиццы на дом (система анализа)»\n")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(14); r.font.color.rgb = RGBColor(30, 41, 59)
    r = p_title.add_run("(Программный комплекс «NewFastPizza»)")
    r.font.name = "Times New Roman"; r.font.size = Pt(12); r.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(48)

    # Signatures table
    sig_tbl = doc.add_table(rows=1, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.autofit = False
    sig_tbl.rows[0].cells[0].width = Inches(3.2)
    sig_tbl.rows[0].cells[1].width = Inches(3.2)
    
    p_left = sig_tbl.rows[0].cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_left.paragraph_format.line_spacing = 1.15
    r = p_left.add_run("Выполнили студенты:\n")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(11)
    r = p_left.add_run("____________________ Ахметов Д.Д.\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)
    r = p_left.add_run("____________________ Зюзин Н.М.\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)
    r = p_left.add_run("____________________ Климов М.Д.\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)

    p_right = sig_tbl.rows[0].cells[1].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_right.paragraph_format.line_spacing = 1.15
    r = p_right.add_run("Проверил:\n")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(11)
    r = p_right.add_run("доцент каф. МОиСТ,\nканд. техн. наук\n")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)
    r = p_right.add_run("_________________ И.Г. Лебедев\n\n«___» ____________ 2026 г.")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(40)

    p_bot = doc.add_paragraph()
    p_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bot.paragraph_format.space_before = Pt(0)
    p_bot.paragraph_format.space_after = Pt(0)
    r = p_bot.add_run("Нижний Новгород\n2026")
    r.font.name = "Times New Roman"; r.font.size = Pt(11)

    doc.add_page_break()

    # =========================================================================
    # СОДЕРЖАНИЕ
    # =========================================================================
    p_toc_title = doc.add_paragraph()
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_toc_title.paragraph_format.space_before = Pt(12)
    p_toc_title.paragraph_format.space_after = Pt(12)
    r = p_toc_title.add_run("Содержание")
    r.bold = True; r.font.name = "Times New Roman"; r.font.size = Pt(14)

    toc_items = [
        ("1. История проекта", "2"),
        ("2. Цели дизайна", "4"),
        ("    2.1. Требования пользователя", "4"),
        ("    2.2. Системные требования", "7"),
        ("    2.3. Сценарии использования", "10"),
        ("3. Исключенные возможности и неподдерживаемые сценарии", "16"),
        ("4. Предположения и зависимости", "18"),
        ("5. Проект решения", "20"),
        ("    5.1. Концептуальный проект", "20"),
        ("    5.2. Логический проект", "23"),
        ("    5.3. Физический проект", "28"),
        ("6. Требования к инсталляции и деинсталляции", "32"),
    ]
    for title, page in toc_items:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(2)
        p_t.paragraph_format.space_after = Pt(2)
        p_t.paragraph_format.line_spacing = 1.15
        r_t = p_t.add_run(title)
        r_t.font.name = "Times New Roman"; r_t.font.size = Pt(11)
        if not title.startswith("    "):
            r_t.bold = True
        
        # Dot leader via text spacing
        dots_len = max(2, 68 - len(title))
        r_dots = p_t.add_run(" " + "." * dots_len + " ")
        r_dots.font.name = "Times New Roman"; r_dots.font.size = Pt(10); r_dots.font.color.rgb = RGBColor(148, 163, 184)
        
        r_p = p_t.add_run(page)
        r_p.font.name = "Times New Roman"; r_p.font.size = Pt(11); r_p.bold = True

    doc.add_page_break()

    # =========================================================================
    # 1. ИСТОРИЯ ПРОЕКТА
    # =========================================================================
    p_h1("1. История проекта")

    p_body(
        "Раздел «История проекта» перечисляет ключевые события в процессе создания решения в хронологической "
        "последовательности, так, как их видит команда разработчиков. Перечисляются важные принятые решения, "
        "вехи и подвехи проекта. Данная информация отражает фактический ход работ, позволяет проанализировать "
        "факторы успешности проекта и сформировать базу инженерного опыта для создания аналогичных аналитических "
        "комплексов в будущем."
    )

    p_body(
        "Проект создания аналитического комплекса «NewFastPizza» стартовал в начале ноября 2026 года в рамках учебного курса "
        "«Объектно-ориентированный анализ и проектирование» («Проектирование и архитектура программных систем») на кафедре МОиСТ "
        "ННГУ им. Н.И. Лобачевского. Разработка осуществлялась студенческой проектной группой в составе 3 человек: "
        "Ахметов Д.Д. (Team Lead / ML-инженер), Зюзин Н.М. (Backend-разработчик / DevOps), Климов М.Д. (Frontend-разработчик / QA). "
        "Общая трудоемкость проекта была рассчитана на 8 рабочих недель (360 человеко-часов). Все работы выполнялись строго "
        "по этапам в соответствии с установленным графиком."
    )

    p_body(
        "В таблице 1.1 приведена хронология прохождения контрольных вех и подвех с указанием плановых и фактических сроков "
        "выполнения, статуса готовности и ключевых инженерных решений, принятых командой."
    )

    history_headers = ["Веха / Подвеха", "Плановый срок", "Факт. сдача", "Статус", "Ключевые результаты и принятые решения"]
    history_data = [
        [
            "Веха 1: Проектирование и ТЗ\nПодвеха 1.1: Концепция",
            "01.11–06.11.2026",
            "06.11.2026",
            "Сдано вовремя",
            "Сбор требований стейкхолдеров сети «NewFastPizza». Формулирование видения, рамок и выгод. Утвержден документ «Концепция проекта»."
        ],
        [
            "Подвеха 1.2: Архитектура и план",
            "07.11–11.11.2026",
            "11.11.2026",
            "Сдано вовремя",
            "Построение матрицы компромиссов, распределение ролей, расчет сметы (247 200 руб.). Утвержден документ «Структура проекта»."
        ],
        [
            "Подвеха 1.3: ER-модель и API",
            "12.11–14.11.2026",
            "14.11.2026",
            "Сдано вовремя",
            "Проектирование 5 базовых реляционных таблиц PostgreSQL. Выбор REST API спецификации (FastAPI / OpenAPI) вместо GraphQL."
        ],
        [
            "Подвеха 1.4: UI-макеты и ТЗ",
            "14.11–16.11.2026",
            "15.11.2026",
            "С опережением\nна 1 день",
            "Разработка интерактивных прототипов веб-дашборда. Утверждение запрета на написание кода до полной защиты диаграммы классов."
        ],
        [
            "Веха 2: Ядро и алгоритмы\nПодвеха 2.1: База данных и ORM",
            "16.11–20.11.2026",
            "19.11.2026",
            "С опережением\nна 1 день",
            "Настройка PostgreSQL 16 в Docker. Создание моделей SQLAlchemy 2.0, настройка системы миграций схемы Alembic."
        ],
        [
            "Подвеха 2.2: Генератор данных",
            "20.11–23.11.2026",
            "22.11.2026",
            "Сдано вовремя",
            "Реализация скрипта генерации демо-данных (7 филиалов в Н. Новгороде, 20 курьеров, тысячи заказов с координатами и датами)."
        ],
        [
            "Подвеха 2.3: Алгоритм маршрутов",
            "23.11–27.11.2026",
            "26.11.2026",
            "Сдано вовремя",
            "Интеграция дорожного графа Нижнего Новгорода через NetworkX. Алгоритм Дейкстры с учетом весов пробок и очередей филиалов."
        ],
        [
            "Подвеха 2.4: Модуль K-Means",
            "27.11–30.11.2026",
            "29.11.2026",
            "Сдано вовремя",
            "Пайплайн кластеризации спроса на Scikit-learn: определение центроидов, сопоставление с филиалами, выработка рекомендаций."
        ],
        [
            "Веха 3: Веб-интерфейс\nПодвеха 3.1: Верстка дашборда",
            "01.12–06.12.2026",
            "05.12.2026",
            "Сдано вовремя",
            "Компонентная верстка на Vue 3 (Composition API), TypeScript, Tailwind CSS. Интеграция графиков динамики и KPI на Chart.js."
        ],
        [
            "Подвеха 3.2: Картография и Folium",
            "06.12–10.12.2026",
            "09.12.2026",
            "Сдано вовремя",
            "Интеграция интерактивных карт Leaflet/Folium: слои тепловых карт спроса (Heatmap), полилинии маршрутов и маркеры филиалов."
        ],
        [
            "Подвеха 3.3: Интеграция и RBAC",
            "10.12–15.12.2026",
            "13.12.2026",
            "С опережением\nна 2 дня",
            "Связка фронтенда и бэкенда по REST API. Авторизация JWT (Bearer), Pinia store, фильтрация маршрутов по ролям пользователей."
        ],
        [
            "Веха 4: Тесты и деплой\nПодвеха 4.1: Комплексные тесты",
            "16.12–21.12.2026",
            "20.12.2026",
            "Сдано вовремя",
            "Разработка 181 теста Pytest (модульные, интеграционные, ML-пайплайн). Достигнуто 100% прохождение тестов на SQLite in-memory."
        ],
        [
            "Подвеха 4.2: Контейнеризация",
            "21.12–24.12.2026",
            "23.12.2026",
            "Сдано вовремя",
            "Сборка легковесного Docker-образа API (python:3.12-slim), настройка compose.yaml с именованными томами postgres_data и app_data."
        ],
        [
            "Подвеха 4.3: CI/CD и VPS",
            "24.12–26.12.2026",
            "25.12.2026",
            "С опережением\nна 1 день",
            "Настройка GitHub Actions (линтер Ruff, pytest, smoke-test.sh). Автоматизированный деплой на тестовый сервер Ubuntu 22.04 LTS."
        ],
        [
            "Подвеха 4.4: Защита проекта",
            "26.12–28.12.2026",
            "27.12.2026",
            "Сдано вовремя",
            "Финальная демонстрация системы заказчику. Выполнение всех критериев приемки. Подготовка комплекта документации."
        ]
    ]
    p_table(history_headers, history_data, [1.4, 0.9, 0.8, 1.0, 2.3])

    p_h2("Ключевые проектные решения, определившие успех системы:")
    p_bullet("1. Отказ от микросервисов в пользу модульного монолита:", 
             "На этапе Вехи 1 команда осознанно отказалась от сложной распределенной микросервисной инфраструктуры. В условиях жестких сроков монолитная слоистая архитектура (Layered Architecture) на базе FastAPI позволила обеспечить высокую скорость разработки и изоляцию модулей без сетевых накладных расходов.")
    p_bullet("2. Предварительный расчет и кэширование дорожного графа в RAM:", 
             "Обработка полного графа дорог города-миллионника (Нижний Новгород) в реальном времени через внешние API приводила к задержкам до 8-10 секунд. Решение упаковать граф в сериализованный формат Pickle (data/nn_graph.pkl) и загружать его в оперативную память сервера при старте приложения снизило время отклика API маршрутизации до 45-80 мс.")
    p_bullet("3. Строгая дисциплина проектирования (No Code before UML):", 
             "Соблюдение основополагающего правила курса — не начинать реализацию до защиты концептуальной архитектуры и подробной диаграммы классов — позволило избежать переделок схемы БД и рефакторинга интерфейсов в процессе разработки.")

    # =========================================================================
    # 2. ЦЕЛИ ДИЗАЙНА
    # =========================================================================
    p_h1("2. Цели дизайна")

    p_body(
        "Данный раздел документа обобщает и развивает выполненный ранее анализ требований к разрабатываемой системе. "
        "Формулируются уточненные требования с точки зрения заказчика («NewFastPizza»), конечных пользователей (менеджеров, "
        "руководства сети, водителей-курьеров и системных администраторов), а также требования к аппаратному и программному "
        "окружению. Требования, сформулированные в документе «Концепция проекта», были скорректированы, конкретизированы "
        "и дополнены техническими параметрами с учетом практического опыта, полученного в ходе проектирования и прототипирования."
    )

    # 2.1. Требования пользователя
    p_h2("2.1. Требования пользователя")
    p_body(
        "Раздел «Требования пользователя» перечисляет выявленные требования к системе со стороны различных групп стейкхолдеров. "
        "Система «NewFastPizza» ориентирована на четыре ключевые категории пользователей с четким разделением функциональных обязанностей:"
    )

    p_bullet("1. Менеджеры филиальной сети:", 
             "Сотрудники, отвечающие за оперативное управление заказами, контроль нагрузки на кухни филиалов, анализ локального спроса и планирование курьерских смен.")
    p_bullet("2. Руководство компании (Executive / Владелец бизнеса):", 
             "Лица, принимающие стратегические инвестиционные решения о развитии сети, открытии новых филиалов, оптимизации логистических расходов и оценке рентабельности.")
    p_bullet("3. Водители-курьеры:", 
             "Персонал службы доставки, непосредственно выполняющий рейсы и нуждающийся в четких инструкциях по выбору следующей точки базирования для минимизации холостого пробега.")
    p_bullet("4. Системный администратор (ИТ-специалист):", 
             "Сотрудник, обеспечивающий бесперебойную эксплуатацию программного комплекса, управление учетными записями сотрудников и актуализацию цифровой карты дорог.")

    p_body("Ниже представлены детализированные пользовательские требования:")

    user_reqs_headers = ["ID", "Категория", "Формулировка требования", "Критерий приемки"]
    user_reqs_data = [
        ["ПТ-01", "Менеджер", "Просмотр аналитического дашборда с ключевыми метриками (KPI)", "Отображение общего числа заказов, среднего времени доставки, процента отмен и среднего чека с возможностью фильтрации по датам."],
        ["ПТ-02", "Менеджер", "Мониторинг загрузки и очередей филиалов сети", "Наглядное распределение заказов по 7 филиалам, отображение текущей очереди на кухне, подсветка перегруженных филиалов."],
        ["ПТ-03", "Менеджер", "Анализ зон спроса методом машинного обучения (K-Means)", "Интерактивная карта с визуализацией кластеров концентрации заказов, расчет радиуса разброса и плотности доставки."],
        ["ПТ-04", "Менеджер", "Получение рекомендаций по развитию сети филиалов", "Вывод списка рекомендаций: координаты и обоснование для открытия новой точки в зоне дефицита; выявление филиалов с низким спросом."],
        ["ПТ-05", "Менеджер", "Конструктор отчетов с гибкой фильтрацией", "Выбор диапазона дат, статусов заказов, филиалов; группировка по часам и дням недели; сохранение персональных шаблонов отчетов."],
        ["ПТ-06", "Менеджер", "Экспорт сформированных отчетов в табличный формат CSV", "Выгрузка данных в один клик в формате CSV с кодировкой UTF-8 BOM для корректного открытия в Excel и LibreOffice."],
        ["ПТ-07", "Руководство", "Просмотр сводной стратегической аналитики сети", "Агрегированные графики динамики выручки, темпов роста заказов и суммарного снижения среднего времени доставки по сети."],
        ["ПТ-08", "Руководство", "Доступ к утвержденным системным отчетам", "Просмотр и экспорт стандартных отчетов верхнего уровня без права редактирования служебных технических настроек."],
        ["ПТ-09", "Курьер", "Получение автоматической рекомендации по выбору филиала", "После завершения доставки курьер отправляет координаты и мгновенно получает оптимальный филиал для получения следующего заказа."],
        ["ПТ-10", "Курьер", "Расчет кратчайшего пути с учетом пробок", "Построение маршрута по дорожному графу города с учетом актуальной скорости движения (весов ребер) и минимизации времени в пути."],
        ["ПТ-11", "Курьер", "Отображение маршрута и расчетного времени прибытия (ETA)", "Визуализация полилинии маршрута на карте, вывод точного расстояния (км), времени поездки (мин) и текущей очереди заказов в филиале."],
        ["ПТ-12", "Администратор", "Управление учетными записями и ролями пользователей (RBAC)", "Создание, редактирование, блокировка пользователей, назначение ролей: admin, manager, executive, driver."],
        ["ПТ-13", "Администратор", "Управление цифровым дорожным графом и пробками", "Загрузка обновленных файлов графа (pkl), просмотр метаданных (число узлов/ребер), ручная корректировка коэффициентов пробок на улицах."],
        ["ПТ-14", "Администратор", "Мониторинг технического состояния системы (Health Check)", "Проверка доступности эндпоинта /api/v1/health/ready, контроль подключения к базе данных и статуса графа в оперативной памяти."]
    ]
    p_table(user_reqs_headers, user_reqs_data, [0.6, 1.0, 2.4, 2.4])

    # 2.2. Системные требования
    p_h2("2.2. Системные требования")
    p_body(
        "Системные требования определяют функции, технические характеристики, ограничения и параметры окружения, "
        "необходимые для надежного и эффективного функционирования программного комплекса «NewFastPizza». "
        "Требования разделены на три взаимосвязанные группы: функциональные требования, требования к аппаратному обеспечению "
        "и требования к программному стеку."
    )

    p_h3("Функциональные системные требования:")
    p_bullet("СТ-Ф-01 (Хранение данных):", "Система должна обеспечивать персистентное реляционное хранение данных о заказах (координаты клиента, филиал, курьер, временные отметки, статус), филиалах, курьерах и пользователях в СУБД PostgreSQL 16.")
    p_bullet("СТ-Ф-02 (Агрегация метрик):", "Система должна производить высокопроизводительный расчет KPI (суммарное количество, среднее время доставки от оформления до вручения, средний чек, доля отмен) на стороне сервера с использованием векторизованных библиотек Pandas и NumPy.")
    p_bullet("СТ-Ф-03 (Маршрутизация на графе):", "Серверный модуль должен реализовывать поиск кратчайшего пути на ориентированном взвешенном графе дорог с помощью алгоритма Дейкстры (NetworkX). Время расчета маршрута до всех филиалов не должно превышать 100 мс.")
    p_bullet("СТ-Ф-04 (Многокритериальный выбор филиала):", "Выбор рекомендованного филиала должен производиться на основе комплексной целевой функции: минимизация суммы времени движения курьера и прогнозируемого времени ожидания сборки пиццы на филиале.")
    p_bullet("СТ-Ф-05 (Машинное обучение K-Means):", "Система должна выполнять кластеризацию массива исторических доставок с помощью алгоритма K-Means (Scikit-learn), вычислять координаты центроидов зон спроса и сопоставлять их с радиусами обслуживания действующих филиалов.")
    p_bullet("СТ-Ф-06 (Геопространственные слои):", "Система должна формировать данные для рендеринга тепловых карт плотности заказов (Heatmap), полилиний маршрутов курьеров и маркеров филиалов в форматах GeoJSON / Leaflet.")
    p_bullet("СТ-Ф-07 (REST API и документация):", "Серверная часть должна предоставлять типизированный REST API на базе FastAPI (31 эндпоинт) с валидацией входных/выходных схем через Pydantic v2 и автоматической генерацией интерактивной документации Swagger UI (/docs) и ReDoc (/redoc).")
    p_bullet("СТ-Ф-08 (Безопасность и RBAC):", "Аутентификация должна выполняться по протоколу OAuth2 Password Bearer с генерацией криптографически стойких JWT-токенов (алгоритм HS256, время жизни 60 минут). Пароли пользователей должны храниться исключительно в виде соленых хэшей Bcrypt.")
    p_bullet("СТ-Ф-09 (Потоковый экспорт):", "Экспорт аналитических отчетов должен осуществляться в виде потоковой передачи данных (StreamingResponse) в формате CSV с разделителем-запятой и обязательным добавлением UTF-8 BOM для корректного отображения кириллицы.")
    p_bullet("СТ-Ф-10 (Генерация тестовых данных):", "В состав комплекса должен входить программный генератор реалистичных синтетических заказов, филиалов и профилей курьеров для нагрузочного тестирования и демонстрации возможностей дашборда.")

    p_h3("Требования к аппаратному обеспечению (Hardware Requirements):")
    p_body(
        "Поскольку серверная часть выполняет ресурсоемкие алгоритмы на графах и операции машинного обучения, "
        "к аппаратному обеспечению предъявляются повышенные требования, приведенные в таблице 2.1."
    )

    hw_headers = ["Узел системы", "Минимальные требования", "Рекомендуемые требования", "Техническое обоснование"]
    hw_data = [
        [
            "Серверная часть\n(Production VPS /\nХост-машина)",
            "• Процессор: 4 vCPU (x86-64)\n• RAM: 8 ГБ\n• Диск: 40 ГБ SSD\n• Сеть: 100 Мбит/с",
            "• Процессор: 8 vCPU (x86-64 с AVX2)\n• RAM: 16 ГБ\n• Диск: 80 ГБ NVMe SSD\n• Сеть: 1 Гбит/с",
            "Граф дорог Нижнего Новгорода занимает ~1.5 ГБ RAM в виде объекта NetworkX DiGraph. Вычисления K-Means и агрегации Pandas требуют буфера памяти. Многоядерный CPU обеспечивает параллельную обработку запросов курьеров."
        ],
        [
            "Рабочее место\nменеджера / руководства\n(Клиентский ПК)",
            "• Процессор: 2 ядра 2.0 ГГц\n• RAM: 4 ГБ\n• Дисплей: 1366×768\n• Сеть: 10 Мбит/с",
            "• Процессор: 4+ ядра 2.5 ГГц\n• RAM: 8+ ГБ\n• Дисплей: 1920×1080 (Full HD)\n• Сеть: 50+ Мбит/с",
            "Клиентское SPA-приложение исполняется в браузере. Отображение векторных интерактивных карт (Leaflet) и графиков Chart.js требует ресурсов GPU/CPU для плавной анимации."
        ],
        [
            "Мобильное устройство\nводителя-курьера",
            "• Экран: от 5.0 дюймов\n• RAM: 2 ГБ\n• Модем: 3G / LTE\n• GPS/ГЛОНАСС модуль",
            "• Экран: от 6.0 дюймов (AMOLED)\n• RAM: 4 ГБ\n• Модем: 4G LTE / 5G\n• GPS/ГЛОНАСС/Galileo",
            "Необходимо стабильное позиционирование курьера на карте и быстрое обновление маршрутов при передвижении по городу."
        ]
    ]
    p_table(hw_headers, hw_data, [1.5, 1.6, 1.6, 1.7])

    p_h3("Требования к программному обеспечению (Software Requirements):")
    p_body(
        "Все компоненты системы построены на базе свободного программного обеспечения с открытым исходным кодом (Open Source). "
        "В таблице 2.2 зафиксирован стек технологий и версии библиотек, используемые в проекте."
    )

    sw_headers = ["Уровень системы", "Технология / Библиотека", "Версия", "Назначение в проекте"]
    sw_data = [
        ["Серверная ОС", "Ubuntu Linux", "22.04 LTS", "Базовая операционная система хоста и контейнеров."],
        ["Контейнеризация", "Docker & Docker Compose", "24.0+ / 2.20+", "Изоляция окружения, сборка образов и оркестрация стека."],
        ["Среда исполнения", "Python (строго 3.12)", "3.12.7", "Интерпретатор backend-сервисов (гарантия совместимости wheels)."],
        ["Веб-фреймворк", "FastAPI", "0.115+", "Асинхронный REST API фреймворк, внедрение зависимостей, OpenAPI."],
        ["ASGI-сервер", "Uvicorn", "0.34+", "Высокопроизводительный сервер запуска FastAPI приложений."],
        ["СУБД", "PostgreSQL", "16-alpine", "Основное реляционное хранилище транзакционных данных."],
        ["ORM и миграции", "SQLAlchemy / Alembic", "2.0+ / 1.14+", "Объектно-реляционное отображение моделей и версионирование БД."],
        ["Обработка данных", "Pandas / NumPy", "2.2+ / 2.1+", "Векторизованные расчеты KPI, агрегация метрик заказов."],
        ["Машинное обучение", "Scikit-learn", "1.6+", "Реализация алгоритма кластеризации K-Means для зон спроса."],
        ["Теория графов", "NetworkX / OSMnx", "3.4+ / 2.0+", "Представление дорожной сети города, расчет алгоритма Дейкстры."],
        ["Безопасность", "PyJWT / Passlib (Bcrypt)", "2.10+ / 4.0+", "Генерация JWT-токенов, безопасное хэширование паролей."],
        ["Клиентский фреймворк", "Vue 3 (Composition API)", "3.5+", "Реактивный интерфейс веб-дашборда."],
        ["Язык фронтенда", "TypeScript", "5.6+", "Строгая типизация данных DTO, моделей и компонентов интерфейса."],
        ["Сборщик фронтенда", "Vite", "5.4+", "Быстрая разработка и оптимизированная сборка продакшен-бандла."],
        ["Стейт-менеджер", "Pinia", "2.2+", "Централизованное хранилище состояния пользователя и сессий."],
        ["Визуализация данных", "Chart.js / Lucide", "4.4+ / 0.45+", "Построение графиков KPI, круговых диаграмм, пиктограмм интерфейса."],
        ["Картография", "Leaflet / Folium", "1.9+ / 0.17+", "Рендеринг интерактивных карт, тепловых слоев и полилиний дорог."],
        ["Стилизация", "Tailwind CSS", "3.4+", "Утилитарный CSS-фреймворк для адаптивной адаптации экранов."],
        ["Тестирование", "Pytest / HTTPX", "8.3+ / 0.28+", "Фреймворк для запуска 181 модульного и интеграционного теста."],
        ["Качество кода", "Ruff", "0.8+", "Сверхбыстрый статический анализатор и форматировщик Python-кода."]
    ]
    p_table(sw_headers, sw_data, [1.3, 1.6, 0.9, 2.6])

    # 2.3. Сценарии использования
    p_h2("2.3. Сценарии использования")
    p_body(
        "Сценарии использования (Use Cases) определяют последовательность взаимодействия акторов (пользователей и внешних систем) "
        "с программным комплексом «NewFastPizza» для достижения конкретных прикладных целей. "
        "На рисунке 2.1 приведена диаграмма прецедентов использования (UML Use Case Diagram), утвержденная в первом документе "
        "«Концепция проекта». Диаграмма отражает ролевое разделение функций между менеджером, руководителем сети, водителем "
        "и администратором системы."
    )

    p_image("docs/diagrams/use_case_diagram.png", "Рис. 2.1. UML Диаграмма прецедентов использования (Use Case Diagram)", width_in=5.8)

    p_body(
        "На диаграмме представлены четыре ключевых актора и соответствующие прецеденты использования:\n"
        "• Менеджер: «Просмотр аналитических показателей», «Формирование и экспорт отчетов», «Получение рекомендаций по открытию или закрытию филиалов».\n"
        "• Руководитель сети: «Просмотр аналитических показателей», «Экспорт сводных отчетов».\n"
        "• Водитель-курьер: «Получение рекомендаций по выбору следующего филиала» (взаимодействие через клиентское приложение).\n"
        "• Администратор системы: «Управление пользователями и правами доступа», «Обновление графа дорог и управление пробками»."
    )

    p_body(
        "В соответствии с требованиями курса, наиболее важные и алгоритмически сложные варианты использования детализируются "
        "с помощью диаграмм деятельности (UML Activity Diagrams / блок-схем). Каждая диаграмма отражает логику технической "
        "реализации прецедента внутри системы, включая обработку данных, принятие решений и взаимодействие между слоями архитектуры."
    )

    # Activity 1
    p_h3("Детализация прецедента 1: «Расчет и выдача рекомендаций курьеру по выбору филиала»")
    p_body(
        "Данный сценарий является ключевым для логистической оптимизации. Цель прецедента — минимизировать суммарное время "
        "простоя курьера за счет одновременного учета дорожной ситуации на графе (пробок) и очередей заказов на филиалах. "
        "На рисунке 2.2 представлена UML Activity-диаграмма технической реализации данного процесса."
    )

    p_image("docs/diagrams/activity_driver_routing.png", "Рис. 2.2. UML Activity-диаграмма: «Расчет и выдача рекомендаций курьеру по выбору филиала»", width_in=5.2)

    p_body(
        "Последовательность выполнения операций сценария маршрутизации курьера:\n"
        "1. Инициация запроса: После доставки пиццы клиенту курьер запрашивает следующий маршрут через клиентское приложение. "
        "Приложение считывает текущие GPS-координаты устройства (lat, lon) и отправляет запрос POST /api/v1/recommendations/driver-branch.\n"
        "2. Аутентификация и валидация: API-шлюз проверяет Bearer JWT-токен курьера и валидирует входные координаты через Pydantic-схему DriverLocationDTO.\n"
        "3. Проверка состояния графа в RAM: Сервис GraphService проверяет, загружен ли дорожный граф Нижнего Новгорода в оперативную память. "
        "В случае штатной работы граф уже находится в памяти; при первом обращении происходит мгновенная десериализация из кэша data/nn_graph.pkl.\n"
        "4. Привязка к дорожной сети: Алгоритм находит ближайший узел графа дорог к координатам курьера с использованием пространственного дерева поиска.\n"
        "5. Выборка активных филиалов: Слой BranchRepository запрашивает из БД PostgreSQL список всех работающих филиалов и размер их текущей очереди заказов (current_queue).\n"
        "6. Расчет кратчайших путей (Дейкстра): Для каждого филиала алгоритм Дейкстры (NetworkX) находит кратчайший путь по взвешенным ребрам графа "
        "(вес ребра отражает время движения в минутах с учетом текущей загруженности дорог).\n"
        "7. Многокритериальный скоринг филиалов: Сервис RoutingAdapter вычисляет итоговый оценочный балл для каждого филиала по формуле:\n"
        "   Score = T_travel + W_queue * current_queue,\n"
        "   где T_travel — расчетное время поездки до филиала (мин), current_queue — количество заказов, ожидающих доставки, "
        "W_queue — весовой коэффициент полезности очереди. Выбирается филиал с наилучшим показателем минимизации ожидания.\n"
        "8. Формирование ответа и рендеринг: Сервер возвращает DTO с идентификатором филиала, его названием, ETA прибытия, "
        "дистанцией в километрах и массивом координат полилинии маршрута. Интерфейс курьера отображает маршрут на карте."
    )

    # Activity 2
    p_h3("Детализация прецедента 2: «Анализ спроса и рекомендации по развитию сети филиалов (K-Means)»")
    p_body(
        "Сценарий предназначен для менеджмента и руководства компании. Цель — выявить пространственные закономерности "
        "распределения заказов по районам города и сгенерировать математически обоснованные рекомендации по открытию новых "
        "точек или оптимизации действующих. На рисунке 2.3 представлена соответствующая UML Activity-диаграмма."
    )

    p_image("docs/diagrams/activity_demand_kmeans.png", "Рис. 2.3. UML Activity-диаграмма: «Анализ спроса и рекомендации по развитию сети филиалов (K-Means)»", width_in=5.2)

    p_body(
        "Техническая последовательность выполнения пространственного анализа спроса:\n"
        "1. Задание параметров анализа: Менеджер в дашборде выбирает временной интервал (например, за последний месяц), "
        "желаемое число кластеров K (по умолчанию K=7) и пороговый радиус обслуживания филиала (обычно 2.5–3.0 км).\n"
        "2. Выборка геоданных: Сервис DemandAdapter запрашивает из БД массив выполненных заказов за указанный период (client_lat, client_lon).\n"
        "3. Проверка объема данных: Проверяется условие N >= K (число заказов должно превосходить число кластеров). При недостатке данных возвращается предупреждение.\n"
        "4. Кластеризация K-Means: Модуль машинного обучения на базе Scikit-learn производит итеративную кластеризацию координат. "
        "Вычисляются координаты K центроидов спроса и средний радиус разброса точек вокруг каждого центроида.\n"
        "5. Пространственный анализ покрытия: Для каждого центроида вычисляется евклидово расстояние до ближайшего действующего филиала сети «NewFastPizza».\n"
        "6. Классификация рекомендаций:\n"
        "   • Если в радиусе центроида высокая плотность заказов, но ближайший филиал находится дальше допустимого радиуса -> формируется рекомендация «ОТКРЫТЬ ФИЛИАЛ» в точке центроида;\n"
        "   • Если вокруг действующего филиала наблюдается крайне низкая концентрация заказов -> формируется рекомендация «ЗАКРЫТЬ/ОПТИМИЗИРОВАТЬ ТОЧКУ»;\n"
        "   • Если спрос филиала сбалансирован -> точка получает статус «Оптимальное покрытие».\n"
        "7. Визуализация: Генерируется слой тепловой карты (Heatmap) и карточки рекомендаций, отображаемые на экране менеджера."
    )

    # Activity 3
    p_h3("Детализация прецедента 3: «Формирование и экспорт аналитического отчета»")
    p_body(
        "Сценарий обеспечивает потребности персонала в получении детализированных срезов операционной деятельности "
        "и выгрузке данных для бухгалтерского и управленческого учета. На рисунке 2.4 приведена UML Activity-диаграмма."
    )

    p_image("docs/diagrams/activity_report_export.png", "Рис. 2.4. UML Activity-диаграмма: «Формирование и экспорт аналитического отчета»", width_in=5.2)

    p_body(
        "Последовательность выполнения операций формирования и экспорта отчетов:\n"
        "1. Выбор шаблона и параметров: Менеджер выбирает системный шаблон (сводка по филиалам, статистика курьеров, распределение по часам) "
        "или создает пользовательский фильтр (диапазон дат, филиалы, статусы заказов).\n"
        "2. Выполнение запроса: Отправляется POST-запрос к /api/v1/reports/custom. Авторизационный фильтр проверяет права пользователя.\n"
        "3. Извлечение данных: OrderRepository выполняет оптимизированный SQL-запрос к таблице orders с использованием составных индексов.\n"
        "4. Агрегация метрик: Сервис ReportService загружает строки в Pandas DataFrame и вычисляет агрегированные показатели: "
        "суммарное число заказов, среднее время доставки, процент успешных доставок, выручку.\n"
        "5. Отображение в браузере: Сводные данные передаются клиенту в формате JSON и отрисовываются в виде интерактивной таблицы и графиков.\n"
        "6. Потоковый экспорт в CSV: При нажатии кнопки «Экспорт в CSV» сервер преобразует DataFrame в текстовый поток CSV "
        "с добавлением BOM (Byte Order Mark, \\xef\\xbb\\xbf) и отправляет файл с заголовком Content-Disposition: attachment; filename=\"report.csv\"."
    )

    # =========================================================================
    # 3. ИСКЛЮЧЕННЫЕ ВОЗМОЖНОСТИ И НЕПОДДЕРЖИВАЕМЫЕ СЦЕНАРИИ
    # =========================================================================
    p_h1("3. Исключенные возможности и неподдерживаемые сценарии")

    p_body(
        "В данном разделе зафиксирован перечень требований и функциональных возможностей, которые осознанно не вошли "
        "в текущий релиз программного комплекса «NewFastPizza». Четкое определение границ решения (Scope) и документирование "
        "неподдерживаемых сценариев имеет критическую важность для исключения разночтений между командой разработчиков, "
        "заказчиком и преподавателем. Все решения об исключении возможностей базируются на проектной матрице компромиссов, "
        "жестких временных рамках (ноябрь–декабрь 2026 г.) и ориентации проекта на создание аналитического ядра."
    )

    excluded_headers = ["№", "Исключенная возможность / сценарий", "Причина и обоснование исключения", "Условия реализации в будущем (Релиз 2.0)"]
    excluded_data = [
        [
            "1",
            "Разработка нативных мобильных приложений для iOS и Android",
            "Фокус проекта сосредоточен на backend-аналитике (ML, графы) и веб-дашборде. Разработка двух нативных мобильных клиентов и прохождение модерации в App Store / Google Play требует отдельных компетенций и удваивает бюджет. Для курьеров предоставлен адаптивный веб-интерфейс и REST API.",
            "Релиз 2.0 (весна 2027 г.) при выделении отдельной мобильной команды разработчиков. Серверный REST API уже полностью готов для интеграции."
        ],
        [
            "2",
            "Онлайн-интеграция с коммерческими картографическими API (Яндекс.Карты, Google Maps)",
            "Коммерческие картографические сервисы тарифицируются помесячно и имеют жесткие лимиты на число геокодирований и маршрутизаций. Это нарушает академическое требование о нулевом бюджете и использовании исключительно Open Source. Вместо них используется локальный граф OpenStreetMap.",
            "Возможно подключение корпоративного API Яндекс.Карт при наличии коммерческого контракта и финансирования заказчика."
        ],
        [
            "3",
            "Автоматическое закрытие или открытие филиалов без участия человека",
            "Система проектируется как система поддержки принятия решений (DSS). Закрытие филиала влечет юридические, арендные и кадровые последствия. Алгоритм K-Means формирует объективные рекомендации, но финальное решение принимает руководство компании.",
            "Полная автоматизация не планируется из соображений безопасности бизнеса. В релизе 2.0 возможна интеграция с системой документооборота."
        ],
        [
            "4",
            "Глубокие нейросетевые модели прогнозирования спроса (LSTM, Deep Learning)",
            "Обучение рекуррентных нейронных сетей требует непрерывных исторических данных за несколько лет, вычислительных кластеров с GPU и длительного тюнинга. Для текущих масштабов сети (до 20 филиалов) алгоритм K-Means обеспечивает 100% интерпретируемость и точность при минимальных ресурсах.",
            "Релиз 2.1 при накоплении более 500 000 реальных транзакций заказов и развертывании GPU-сервера."
        ],
        [
            "5",
            "Интеграция с эквайрингом, онлайн-кассами (ФЗ-54) и бухгалтерскими ERP (1C)",
            "Команда реализует 3-ю и 4-ю части корпоративной системы (модули сбора статистики и аналитических рекомендаций). Модули клиентской оплаты и фискализации чеков относятся к 1-й части общей системы («Клиентский сервис»).",
            "Интеграция через защищенный шлюз Webhooks при объединении всех четырех частей корпоративной платформы."
        ],
        [
            "6",
            "Распределенная шина сообщений (Apache Kafka, RabbitMQ) и микросервисы",
            "На текущем объеме трафика (до 100 запросов в секунду) внедрение Kafka привело бы к избыточной сложности (overengineering), росту задержек сериализации и усложнению отладки. Модульный монолит полностью удовлетворяет требованиям по производительности.",
            "Миграция на брокер сообщений планируется при масштабировании сети NewFastPizza на другие регионы России (более 100 филиалов)."
        ]
    ]
    p_table(excluded_headers, excluded_data, [0.4, 1.8, 2.4, 1.8])

    # =========================================================================
    # 4. ПРЕДПОЛОЖЕНИЯ И ЗАВИСИМОСТИ
    # =========================================================================
    p_h1("4. Предположения и зависимости")

    p_body(
        "Раздел «Предположения и зависимости» перечисляет допущения, заложенные в основу архитектурного проекта "
        "и алгоритмических решений, а также определяет критические внешние зависимости, необходимые для стабильной "
        "работы функциональной спецификации и развернутого программного обеспечения."
    )

    p_h2("Проектные предположения (Assumptions):")
    p_bullet("1. Полнота и связность дорожного графа:", 
             "Предполагается, что дорожная сеть города представлена связным ориентированным взвешенным графом. Узлы соответствуют перекресткам, а ориентированные ребра — направлениям движения с положительными весами времени проезда. Изолированные тупиковые узлы исключаются на этапе предварительной фильтрации графа.")
    p_bullet("2. Дисциплина следования курьеров рекомендациям:", 
             "Предполагается, что после завершения доставки курьеры направляются именно в тот филиал, который указан системой, и следуют по предложенному кратчайшему маршруту, что обеспечивает расчетное снижение общего простоя сети на 15%.")
    p_bullet("3. Квазистатичность дорожной обстановки на шаге расчета:", 
             "Веса дорожных ребер (коэффициенты пробок) считаются постоянными на протяжении интервала расчета отдельного маршрута (несколько десятков миллисекунд). Динамическое изменение дорожной ситуации учитывается при периодическом обновлении графа администратором.")
    p_bullet("4. Достоверность координат входных заказов:", 
             "Предполагается, что координаты доставок поступают в систему в стандартной геодезической системе координат WGS 84 (EPSG:4326) и находятся строго в пределах географических границ города Нижнего Новгорода.")

    p_h2("Внешние зависимости (Dependencies):")
    p_bullet("1. Зависимость от среды исполнения Python 3.12:", 
             "Все математические и серверные библиотеки жестко скомпилированы под версию Python 3.12. Использование других версий (например, 3.13) не поддерживается из-за отсутствия скомпилированных бинарных колес для библиотек osmnx, scikit-learn и pandas.")
    p_bullet("2. Зависимость от СУБД PostgreSQL 16:", 
             "База данных является единственным источником правды для реляционных сущностей, внешних ключей и исторических транзакций. Работа без доступной СУБД невозможна (при отказе БД эндпоинт /health/ready возвращает статус ошибки 503).")
    p_bullet("3. Зависимость от инфраструктуры Docker и Docker Compose:", 
             "Корректное развертывание и межсервисное взаимодействие backend-сервера и базы данных обеспечивается средствами внутренней виртуальной сети Docker (`pizza-network`).")
    p_bullet("4. Зависимость от наличия кэша графа дорог (`data/nn_graph.pkl`):", 
             "Для быстрого старта алгоритмов маршрутизации необходим локальный файл графа дорог. При его отсутствии в системе предусмотрен fallback-генератор синтетического графа для непрерывности прогона юнит-тестов.")
    p_bullet("5. Зависимость от поддержки веб-стандартов браузерами пользователей:", 
             "Пользовательский интерфейс использует возможности современных веб-браузеров (Fetch API, WebAssembly, CSS Grid/Flexbox, ES2022). Работа в устаревших браузерах (Internet Explorer) не поддерживается.")

    # =========================================================================
    # 5. ПРОЕКТ РЕШЕНИЯ
    # =========================================================================
    p_h1("5. Проект решения")

    p_body(
        "Раздел «Проект решения» является ключевой частью функциональной спецификации и обобщает проектные решения, "
        "созданные в процессе объектно-ориентированного анализа. Раздел структурирован по трем каноническим уровням "
        "проектирования: концептуальный проект (архитектура в терминах пользователя и бизнес-процессов), "
        "логический проект (объектная абстрактная модель, иерархия классов, атрибуты и связи) и физический проект "
        "(проекция решения на физические узлы, контейнеры и сетевые протоколы)."
    )

    # 5.1. Концептуальный проект
    p_h2("5.1. Концептуальный проект")
    p_body(
        "Концептуальный проект представляет стратегическое описание системы «глазами пользователя» без избыточного "
        "программистского сленга. Назначение системы «NewFastPizza» — трансформировать хаотичную интуитивную логику доставки "
        "в прозрачный, математически управляемый процесс. "
        "На рисунке 5.1 представлена концептуальная схема архитектуры решения, демонстрирующая сквозное взаимодействие "
        "всех участников процесса через центральный аналитический комплекс."
    )

    p_image("docs/diagrams/conceptual_architecture.png", "Рис. 5.1. Концептуальная архитектура аналитического комплекса «NewFastPizza»", width_in=6.2)

    p_body(
        "Архитектура концептуально разделена на четыре понятных для бизнеса уровня:\n"
        "1. Уровень пользователей и стейкхолдеров: Менеджеры филиалов получают удобный инструмент контроля очередей и аналитики; "
        "руководство компании оперирует объективными данными для открытия новых филиалов; курьеры получают четкие указания, "
        "куда ехать после доставки; администратор осуществляет централизованный контроль прав доступа.\n"
        "2. Уровень пользовательских интерфейсов: Представлен адаптивным веб-дашбордом (доступен с любых рабочих мест через браузер), "
        "интерактивными картографическими экранами и мобильным шлюзом для взаимодействия с терминалами водителей.\n"
        "3. Уровень центрального аналитического сервиса: Ядро комплекса, объединяющее пять специализированных модулей: "
        "модуль авторизации и безопасности (контроль прав), аналитический модуль KPI (расчет времени доставки и выручки), "
        "конструктор отчетов (экспорт срезов), интеллектуальную маршрутизацию курьеров (алгоритм Дейкстры с учетом пробок) "
        "и пространственный ML-анализ спроса (кластеризация K-Means).\n"
        "4. Уровень надежного хранения данных: Центральная база данных всех транзакций сети и цифровая карта дорог города."
    )

    p_body(
        "В таблице 5.1 приведен сравнительный анализ альтернативных архитектурных решений, рассмотренных командой "
        "на этапе концептуального проектирования, с обоснованием итогового выбора."
    )

    arch_alt_headers = ["Аспект дизайна", "Рассмотренные альтернативы", "Плюсы и минусы", "Итоговое решение и обоснование"]
    arch_alt_data = [
        [
            "Тип клиентского приложения",
            "1. Десктопный клиент (PyQt/Electron)\n2. Веб-приложение (SPA)",
            "Десктоп требует ручной установки и обновлений на каждом ПК менеджера. Веб-SPA доступно мгновенно по URL через любой браузер без настройки.",
            "Веб-приложение (Vue 3 + TypeScript). Обеспечивает нулевую стоимость внедрения на рабочих местах и легкую кроссплатформенность."
        ],
        [
            "Архитектура бэкенда",
            "1. Микросервисы (FastAPI + gRPC)\n2. Модульный монолит (FastAPI Layered)",
            "Микросервисы дают независимый деплой, но требуют сетевой шины и сложного мониторинга. Модульный монолит прост в отладке и имеет нулевые задержки вызова сервисов.",
            "Модульный монолит. Оптимален по критерию «сложность/сроки/надежность» для команды из 3 разработчиков."
        ],
        [
            "Алгоритм анализа спроса",
            "1. DBSCAN (плотностная кластеризация)\n2. K-Means (кластеризация центроидов)",
            "DBSCAN чувствителен к подбору радиуса epsilon и оставляет «шум». K-Means гарантированно выделяет K компактных зон с точными центроидами для филиалов.",
            "Алгоритм K-Means (Scikit-learn). Дает математически интерпретируемые координаты для рекомендаций открытия филиалов."
        ],
        [
            "Способ маршрутизации",
            "1. Внешний геокодер (OSRM API)\n2. Локальный граф NetworkX в RAM",
            "Внешний сервис зависит от доступности стороннего сервера и сети. Локальный граф в RAM обеспечивает отклик <100 мс и 100% автономность.",
            "Локальный граф NetworkX с сериализацией в pickle. Полная независимость от интернета и максимальная скорость."
        ]
    ]
    p_table(arch_alt_headers, arch_alt_data, [1.1, 1.6, 2.0, 1.7])

    # 5.2. Логический проект
    p_h2("5.2. Логический проект")
    p_body(
        "Логический проект представляет собой абстрактную объектную модель решения, определяющую структурные элементы, "
        "их атрибуты, методы, иерархию и взаимные связи в нотации языка UML. "
        "В соответствии с методологией курса, диаграмма классов является центральным связующим звеном между анализом требований "
        "и программной реализацией."
    )

    p_callout(
        "Главное академическое правило курса: написание программного кода строго запрещено до тех пор, "
        "пока подробная диаграмма классов (UML Class Diagram) не проверена и официально не утверждена "
        "преподавателем (заказчиком). Нарушение данного правила ведет к архитектурному хаосу и аннулированию результатов этапа.",
        title="ОСНОВОПОЛАГАЮЩЕЕ ПРАВИЛО КУРСА (NO CODE BEFORE UML):",
        bg_color="FEF2F2",
        border_color="DC2626"
    )

    p_body(
        "На рисунке 5.2 представлена детальная UML Диаграмма классов системы «NewFastPizza». Модель декомпозирована "
        "на четыре логических пакета: сущности базы данных (ORM Entities), уровень абстракции доступа к данным (Repositories), "
        "уровень сервисов предметной области и машинного обучения (Business Services & ML) и уровень REST-контроллеров API."
    )

    p_image("docs/diagrams/uml_class_diagram.png", "Рис. 5.2. UML Диаграмма классов системы «NewFastPizza» (Class Diagram)", width_in=6.4)

    p_body(
        "Детальное описание классов и структурных элементов логической модели:\n\n"
        "1. Пакет сущностей предметной области (app.db.models):\n"
        "• User: Класс учетной записи пользователя системы. Атрибуты: id (PK, int), username (UK, str), email (UK, str), "
        "password_hash (str), role (UserRole: admin, manager, executive, driver), is_active (bool), created_at (datetime). "
        "Методы: verify_password(plain_password), set_password(raw_password). Связи: один-к-одному с сущностью Driver (0..1), "
        "один-ко-многим с ReportTemplate (1..*).\n"
        "• Branch: Класс филиала пиццерии. Атрибуты: id (PK, int), name (UK, str), lat (float), lon (float), "
        "current_queue (int, ограничение non-negative), is_active (bool), created_at (datetime). Методы: update_queue(delta), "
        "get_coordinates(). Связи: один-ко-многим с Order (1..*), один-ко-многим с Driver (1..*).\n"
        "• Driver: Класс профиля водителя-курьера. Атрибуты: id (PK, int), full_name (str), phone (str), "
        "branch_id (FK, int, ссылка на Branch), user_id (FK, UK, int, ссылка на User), is_active (bool). "
        "Методы: assign_branch(branch_id), toggle_active(). Связи: один-ко-многим с Order (1..*).\n"
        "• Order: Класс заказа на доставку пиццы. Атрибуты: id (PK, int), external_id (UK, str), client_lat (float), "
        "client_lon (float), branch_id (FK, int), driver_id (FK, int, nullable), order_time (datetime), "
        "delivered_at (datetime, nullable), delivery_time_minutes (int, nullable), status (OrderStatus: Pending, InDelivery, Delivered, Canceled). "
        "Методы: mark_delivered(time).\n"
        "• ReportTemplate: Класс конфигурации аналитического отчета. Атрибуты: id (PK, int), name (str), description (str), "
        "is_system (bool), owner_id (FK, int, nullable — для системных отчетов NULL), config (JSON-конфигурация колонок и фильтров).\n\n"
        "2. Пакет репозиториев доступа к данным (app.repositories):\n"
        "• BaseRepository[T]: Обобщенный интерфейс паттерна Repository. Предоставляет базовые операции абстракции SQLAlchemy: "
        "get_by_id(id), list_all(), create(entity), update(entity), delete(id).\n"
        "• UserRepository: Специфичные запросы к учетным записям: get_by_username(username), get_by_email(email), list_paged(offset, limit).\n"
        "• BranchRepository: Методы выборки филиалов: list_active(), update_queue_val(branch_id, queue), get_with_orders().\n"
        "• OrderRepository: Запросы к массиву заказов: get_orders_in_range(start_date, end_date, branch_id, status), "
        "get_branch_counts(), get_delivery_kpis(), create_batch(orders_list).\n"
        "• DriverRepository: Запросы профилей курьеров: get_by_user_id(user_id), list_active_drivers().\n"
        "• ReportTemplateRepository: Выборка шаблонов: list_for_user(user_id, include_system).\n\n"
        "3. Пакет сервисов бизнес-логики и ML (app.services & app.ml):\n"
        "• AuthService: Сервис аутентификации. Отвечает за проверку учетных данных пользователя, хэширование паролей и выпуск JWT-токенов.\n"
        "• StatsService: Сервис агрегации метрик. Принимает данные от OrderRepository и производит векторизованный расчет KPI с помощью Pandas.\n"
        "• ReportService: Сервис аналитических отчетов. Формирует срезы данных по JSON-шаблонам и реализует потоковый генератор CSV-файлов.\n"
        "• RoutingAdapter & RoutingService: Адаптер и алгоритмическое ядро маршрутизации курьеров на графе дорог NetworkX (поиск Дейкстры, расчет скоринга очередей филиалов).\n"
        "• DemandAdapter & KMeansAnalyzer: Адаптер и модуль машинного обучения для кластеризации точек заказов (Scikit-learn), расчета центроидов и классификации рекомендаций по открытию филиалов.\n"
        "• GraphService: Сервис управления состоянием графа дорог в оперативной памяти (десериализация pickle, обновление весов пробок).\n\n"
        "4. Пакет контроллеров API (app.api.v1.endpoints):\n"
        "Контроллеры осуществляют маршрутизацию HTTP-запросов, валидацию Pydantic DTO схем и инъекцию зависимостей сервисов. "
        "Включает роутеры: AuthRouter, StatsRouter, ReportsRouter, RecommendationsRouter, AdminRouter."
    )

    # 5.3. Физический проект
    p_h2("5.3. Физический проект")
    p_body(
        "Физический проект представляет собой проекцию логической архитектуры на конкретное аппаратное, программное "
        "и сетевое окружение. В этом разделе описывается размещение программных компонентов по физическим узлам, "
        "контейнерам, портам и протоколам передачи данных. "
        "На рисунке 5.3 представлена UML Диаграмма развертывания (Deployment Diagram) целевого релизного комплекса «NewFastPizza»."
    )

    p_image("docs/diagrams/uml_deployment_diagram.png", "Рис. 5.3. UML Диаграмма развертывания (Deployment Diagram)", width_in=6.2)

    p_body(
        "Анализ топологии развертывания и физических компонентов системы:\n\n"
        "1. Физический узел «Рабочая станция менеджера / руководства» (Client Workstation):\n"
        "• Аппаратная платформа: Персональный компьютер или ноутбук под управлением Windows 10/11, macOS или Linux.\n"
        "• Среда исполнения: Современный веб-браузер (Google Chrome, Firefox, Edge, Safari).\n"
        "• Развертываемый артефакт: Клиентский бандл Single Page Application (SPA), скомпилированный Vite (HTML5, JS, CSS). "
        "Исполняет библиотеки реактивности Vue 3, хранилище Pinia, картографию Leaflet и графики Chart.js.\n"
        "• Сетевой канал: Защищенный протокол HTTPS (порт 443 TCP) к серверу VPS.\n\n"
        "2. Физический узел «Мобильный терминал курьера» (Driver Mobile Device):\n"
        "• Аппаратная платформа: Смартфон курьера под управлением Android или iOS с модулем геолокации GPS.\n"
        "• Среда исполнения: Мобильный браузер или клиентское приложение.\n"
        "• Взаимодействие: Передача координат и получение маршрутов по протоколу HTTPS REST JSON.\n\n"
        "3. Физический узел «Целевой сервер: Production VPS Host» (Ubuntu 22.04 LTS):\n"
        "• Аппаратные ресурсы: 4–8 vCPU, 8–16 ГБ RAM, 40+ ГБ SSD, статический внешний IP-адрес.\n"
        "• Уровень хоста (Reverse Proxy): Веб-сервер Nginx слушает внешние порты 80 (HTTP-редирект) и 443 (HTTPS TLS 1.3). "
        "Nginx выполняет терминацию SSL-шифрования, напрямую раздает статические файлы фронтенда из каталога dist/ "
        "и проксирует запросы по маске /api/* во внутреннюю сеть Docker на порт 8000.\n"
        "• Виртуальная сеть Docker (`pizza-network`): Изолированный программный сетевой мост (bridge, подсеть 172.20.0.0/16).\n"
        "• Контейнер `pizza-api` (Execution Environment): Построен на базе официального легковесного образа python:3.12-slim. "
        "Включает ASGI-сервер Uvicorn (порт 8000), приложение FastAPI, библиотеки NumPy, Pandas, Scikit-learn, NetworkX, SQLAlchemy. "
        "Контейнер монтирует постоянный том `app_data` для доступа к файлу дорожного графа data/nn_graph.pkl.\n"
        "• Контейнер `pizza-pg` (Database Engine): Построен на базе postgres:16-alpine. Слушает внутренний порт 5432 TCP, "
        "недоступный из внешней сети интернет. Монтирует постоянный том `postgres_data` в каталог /var/lib/postgresql/data "
        "для обеспечения сохранности данных при перезапусках контейнеров.\n"
        "• Межсервисный протокол: Контейнер pizza-api взаимодействует с pizza-pg по высокоскоростному пулу TCP-соединений "
        "с использованием драйвера psycopg3."
    )

    # =========================================================================
    # 6. ТРЕБОВАНИЯ К ИНСТАЛЛЯЦИИ И ДЕИНСТАЛЛЯЦИИ
    # =========================================================================
    p_h1("6. Требования к инсталляции и деинсталляции")

    p_body(
        "Раздел регламентирует технические требования, пошаговые инструкции и правила проведения процедур инсталляции, "
        "первоначальной инициализации, обновления релизов и деинсталляции программного комплекса «NewFastPizza». "
        "Основным стандартом поставки решения является контейнеризация Docker Compose, что полностью исключает проблемы "
        "несовместимости библиотек и гарантирует идентичность функционирования системы на серверах любой конфигурации."
    )

    p_h2("Требования к процессу инсталляции (Развертывание системы):")
    p_bullet("1. Предварительные требования к хост-машине:", 
             "На сервере под управлением ОС Linux (Ubuntu 22.04 LTS) или локальной машине разработчика должны быть предварительно установлены: Git 2.30+, Docker Engine 24.0+ и плагин Docker Compose v2.20+. Пользователь, производящий установку, должен иметь права суперпользователя (sudo) или входить в системную группу docker.")
    p_bullet("2. Шаг 1: Клонирование дистрибутива из репозитория:", 
             "Выполняется загрузка исходных кодов и конфигурационных файлов команды разработчиков в рабочий каталог:\n"
             "git clone https://github.com/gagashga/Pizza-Analysis-System.git /opt/pizza-analysis\n"
             "cd /opt/pizza-analysis")
    p_bullet("3. Шаг 2: Конфигурирование переменных окружения (.env):", 
             "На основе приложенного эталонного шаблона .env.example создается рабочий файл .env с установкой прав доступа 600:\n"
             "cp .env.example .env && chmod 600 .env\n"
             "В файле обязательно настраиваются: POSTGRES_PASSWORD (пароль к БД), JWT_SECRET_KEY (криптографический ключ подписи токенов), FIRST_ADMIN_USERNAME и FIRST_ADMIN_PASSWORD (логин и пароль администратора).")
    p_bullet("4. Шаг 3: Сборка образов и запуск контейнеров:", 
             "Запуск многоконтейнерного стека в фоновом режиме:\n"
             "docker compose --env-file .env up -d --build\n"
             "Команда автоматически скачивает базовые образы, выполняет компиляцию backend-зависимостей и стартует контейнеры pizza-api и pizza-pg в изолированной сети pizza-network.")
    p_bullet("5. Шаг 4: Применение миграций схемы базы данных:", 
             "Применение актуальных миграций Alembic для автоматического создания таблиц, индексов и ограничений:\n"
             "docker compose exec api alembic -c backend/alembic.ini upgrade head")
    p_bullet("6. Шаг 5: Инициализация демонстрационных данных и администратора:", 
             "Наполнение базы данных 7 филиалами Нижнего Новгорода, 20 профилями курьеров и синтетическими заказами:\n"
             "docker compose exec api python -m backend.app.scripts.seed\n"
             "Скрипт проверяет наличие администратора и создает первичную учетную запись с правами role='admin'.")
    p_bullet("7. Шаг 6: Smoke-тестирование и проверка готовности (Health Check):", 
             "Выполнение проверочного скрипта проверки состояния готовности сервиса:\n"
             "sh scripts/smoke-test.sh\n"
             "Скрипт опрашивает эндпоинт http://localhost:8000/api/v1/health/ready. Ответ {\"status\": \"ok\", \"database\": \"connected\", \"graph\": \"ready\"} свидетельствует об успешном завершении инсталляции.")

    p_h2("Требования к процедуре обновления (Release Update):")
    p_body(
        "Процедура обновления системы на новый релиз занимает менее 2 минут и выполняется без потери накопленных данных:\n"
        "1. Получение изменений из ветки main: git pull origin main\n"
        "2. Пересборка контейнеров приложения: docker compose --env-file .env up -d --build api\n"
        "3. Применение новых миграций БД (если они включены в релиз): docker compose exec api alembic upgrade head\n"
        "Поскольку все транзакции сохраняются в Docker Volume postgres_data, перезапуск контейнеров не затрагивает данные."
    )

    p_h2("Требования к процедуре деинсталляции:")
    p_body(
        "В зависимости от целей вывода системы из эксплуатации предусмотрено два регламента деинсталляции:\n\n"
        "1. Штатная деинсталляция с сохранением данных (Graceful Shutdown):\n"
        "Применяется при временной остановке сервера, переносе на другой хост или регламентном обслуживании оборудования:\n"
        "docker compose down\n"
        "Результат: Контейнеры останавливаются и удаляются, виртуальная сеть закрывается. Именованные тома данных "
        "postgres_data (вся история заказов и филиалов) и app_data (кэш графа) сохраняются на сервере в неизменном виде.\n\n"
        "2. Полная деинсталляция с удалением данных (Purge / Clean Uninstall):\n"
        "Применяется при окончательном выводе сервера из эксплуатации или полной очистке стенда:\n"
        "docker compose down --volumes --rmi all\n"
        "Результат: Полное удаление контейнеров, сетевых мостов, Docker-образов и безвозвратное уничтожение томов данных. "
        "Дополнительно удаляется рабочий каталог проекта: rm -rf /opt/pizza-analysis."
    )

    p_h2("Регламент резервного копирования и восстановления (Disaster Recovery):")
    p_bullet("• Резервное копирование (Backup):", 
             "Перед проведением процедур обновления или деинсталляции администратор обязан создать аварийный дамп базы данных:\n"
             "docker compose exec -T db pg_dump -U pizza pizza > backup_$(date +%Y%m%d).sql\n"
             "и скопировать файл дорожного графа data/nn_graph.pkl на резервный сервер.")
    p_bullet("• Аварийное восстановление (Restore):", 
             "Восстановление данных из дампа выполняется одной командой:\n"
             "cat backup_20261118.sql | docker compose exec -T db psql -U pizza -d pizza")

    # Save final document
    out_dir = "docs"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "3. Функциональная спецификация.docx")
    doc.save(out_path)
    print(f"Document successfully created and saved to: {out_path}")
    return out_path

if __name__ == '__main__':
    create_document()


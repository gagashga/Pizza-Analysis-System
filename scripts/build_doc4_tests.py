# -*- coding: utf-8 -*-
"""
Builder for Document 4: Test Specifications and Scenarios (Спецификации и сценарии тестов)
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

    def p_callout(text, title="ВАЖНОЕ ТРЕБОВАНИЕ К ОРГАНИЗАЦИИ ТЕСТИРОВАНИЯ:", bg_color="FEF3C7", border_color="D97706"):
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

    def p_test_case(test_id, test_name, behavior_in, behavior_op, behavior_out, perf, validity="100%", remarks="—"):
        """Formats a single test case with standard academic elements."""
        p_t = doc.add_paragraph()
        p_t.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_t.paragraph_format.space_before = Pt(10)
        p_t.paragraph_format.space_after = Pt(4)
        p_t.paragraph_format.keep_with_next = True
        r_id = p_t.add_run(f"Тест {test_id}. {test_name}")
        r_id.bold = True
        r_id.font.name = "Times New Roman"
        r_id.font.size = Pt(11)
        r_id.font.color.rgb = RGBColor(15, 23, 42)

        # Expected behavior: Input -> Operation -> Output
        p_b = doc.add_paragraph()
        p_b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_b.paragraph_format.space_before = Pt(1)
        p_b.paragraph_format.space_after = Pt(2)
        p_b.paragraph_format.line_spacing = 1.15
        r_lbl = p_b.add_run("Ожидаемое поведение: ")
        r_lbl.bold = True
        r_lbl.font.name = "Times New Roman"
        r_lbl.font.size = Pt(10.5)
        
        r_in_lbl = p_b.add_run("Входные данные: ")
        r_in_lbl.italic = True
        r_in_lbl.font.name = "Times New Roman"; r_in_lbl.font.size = Pt(10)
        p_b.add_run(behavior_in + " → ").font.name = "Times New Roman"

        r_op_lbl = p_b.add_run("Операция: ")
        r_op_lbl.italic = True
        r_op_lbl.font.name = "Times New Roman"; r_op_lbl.font.size = Pt(10)
        p_b.add_run(behavior_op + " → ").font.name = "Times New Roman"

        r_out_lbl = p_b.add_run("Ожидаемый результат: ")
        r_out_lbl.italic = True
        r_out_lbl.font.name = "Times New Roman"; r_out_lbl.font.size = Pt(10)
        p_b.add_run(behavior_out).font.name = "Times New Roman"

        # Performance, Validity, Remarks in compact bulleted lines
        items = [
            ("Ожидаемая производительность: ", perf),
            ("Ожидаемая достоверность: ", validity),
            ("Замечания: ", remarks)
        ]
        for lbl, val in items:
            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(2)
            p_sub.paragraph_format.line_spacing = 1.15
            r1 = p_sub.add_run(lbl)
            r1.bold = True
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(10)
            r1.font.color.rgb = RGBColor(51, 65, 85)
            r2 = p_sub.add_run(val)
            r2.font.name = "Times New Roman"
            r2.font.size = Pt(10)

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
    r = p_title.add_run("СПЕЦИФИКАЦИИ И СЦЕНАРИИ ТЕСТОВ\n")
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
        ("1. Спецификации на входы", "2"),
        ("2. Спецификации на выходы", "4"),
        ("3. Среда тестирования", "6"),
        ("4. Специальные требования", "8"),
        ("5. Взаимозависимости", "9"),
        ("6. Тесты", "10"),
        ("    6.1. Модульное тестирование (Unit Tests)", "10"),
        ("    6.2. Интеграционное тестирование (Integration Tests)", "26"),
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
        
        dots_len = max(2, 68 - len(title))
        r_dots = p_t.add_run(" " + "." * dots_len + " ")
        r_dots.font.name = "Times New Roman"; r_dots.font.size = Pt(10); r_dots.font.color.rgb = RGBColor(148, 163, 184)
        
        r_p = p_t.add_run(page)
        r_p.font.name = "Times New Roman"; r_p.font.size = Pt(11); r_p.bold = True

    doc.add_page_break()

    # =========================================================================
    # 1. СПЕЦИФИКАЦИИ НА ВХОДЫ
    # =========================================================================
    p_h1("1. Спецификации на входы")

    p_body(
        "Раздел «Спецификации на входы» регламентирует входы, необходимые для прохождения всех тестов разработанного "
        "решения «NewFastPizza». Описание спецификаций наряду с наименованием входа содержит исчерпывающие характеристики: "
        "применяемые тестовые фреймворки, физическое и стендовое оборудование, структуры данных, сетевые протоколы, "
        "базы данных, конфигурационные файлы и программно эмулируемые действия пользователей."
    )

    p_body(
        "В соответствии с инженерной концепцией проекта, ключевой акцент сделан на автоматизированное тестирование. "
        "Для выполнения полного тестового набора используются следующие программные и аппаратные входы:"
    )

    p_bullet("1. Тестовый фреймворк и библиотеки тестирования:", 
             "Основным инструментом запуска тестов выступает фреймворк Pytest (версия 8.3+). Для выполнения асинхронных HTTP-запросов к эндпоинтам FastAPI без поднятия внешнего сетевого сокета используется библиотека HTTPX (AsyncClient, ASGITransport). Для проверки качества и стиля кода применяется статический анализатор Ruff.")
    p_bullet("2. Тестовая база данных (In-Memory Database):", 
             "Для изоляции тестов и достижения максимальной производительности используется легковесная реляционная СУБД SQLite в оперативной памяти (строка подключения: sqlite+aiosqlite:///:memory:). SQLite инициализируется полной схемой SQLAlchemy-моделей (таблицы users, branches, drivers, orders, report_templates) с сохранением внешних ключей и CHECK-ограничений.")
    p_bullet("3. Синтетический взвешенный граф дорог (Фикстура synthetic_graph):", 
             "Для детерминированного тестирования алгоритмов маршрутизации и Дейкстры на вход подается компактный ориентированный мультиграф (NetworkX MultiDiGraph), содержащий 7 узлов и 8 ориентированных ребер с известными весами времени (travel_time) и длины (length). Для нагрузочных тестов используется бинарный кэш реального графа Нижнего Новгорода data/nn_graph.pkl.")
    p_bullet("4. Синтетические наборы данных о заказах и персонале:", 
             "Генератор тестовых данных seed.py и фикстуры создают массивы заказов (идентификаторы ORD-XXXX, координаты доставок в полигоне города, статусы Delivered/Pending/Canceled, время доставки), профили филиалов (с координатами и очередями заказов) и тестовых пользователей с различными ролями.")
    p_bullet("5. Входные файлы конфигураций и весов графа:", 
             "Файлы формата CSV и JSON, передаваемые через Multipart Form Data для тестирования загрузки весов пробок (параметры: u, v, key, weight/multiplier) и описания графов в формате GraphML.")
    p_bullet("6. Физическое оборудование рабочего места инженера по качеству:", 
             "Персональный компьютер (ноутбук) с манипулятором типа «мышь», клавиатурой, сетевым адаптером Ethernet/Wi-Fi и монитором с разрешением не менее 1920×1080 для ручного исследовательского тестирования пользовательского интерфейса в браузере.")

    p_body("В таблице 1.1 систематизированы основные категории входных спецификаций:")

    inputs_headers = ["Категория входа", "Наименование инструмента / ресурса", "Тип / Формат", "Назначение в процессе тестирования"]
    inputs_data = [
        ["Фреймворк", "Pytest 8.3+", "CLI / Python Runner", "Обнаружение, запуск, сборка отчетов и параметризация 181 теста."],
        ["HTTP-клиент", "HTTPX (AsyncClient)", "In-process ASGI client", "Прямая отправка REST-запросов к FastAPI без накладных расходов сети."],
        ["База данных", "SQLite In-Memory", "SQL / RAM storage", "Изолированное хранилище данных, обнуляемое для каждого теста."],
        ["Граф дорог", "NetworkX MultiDiGraph", "Graph in RAM / Pickle", "Математическая модель дорожной сети для проверки алгоритма Дейкстры."],
        ["Данные заказов", "DataFrame / Seed orders", "Pandas DataFrame / Dict", "Массивы транзакций доставок для проверки KPI и K-Means кластеризации."],
        ["Файлы пробок", "weights.csv, graph.graphml", "Text / XML / JSON", "Входные файлы для тестирования загрузки и валидации администратором."],
        ["Аппаратура", "Рабочая станция QA", "x86-64 PC, клавиатура, мышь", "Запуск автотестов, ручная верификация UI-дашборда в браузере."]
    ]
    p_table(inputs_headers, inputs_data, [1.1, 1.8, 1.4, 2.1])

    # =========================================================================
    # 2. СПЕЦИФИКАЦИИ НА ВЫХОДЫ
    # =========================================================================
    p_h1("2. Спецификации на выходы")

    p_body(
        "Раздел «Спецификации на выходы» регламентирует выходы, ожидаемые как результат выполнения тестов для программного "
        "комплекса «NewFastPizza». Выходы включают текстовые консольные логи прохождения тестов, структурированные ответы API, "
        "файлы выгрузки отчетов, сгенерированные интерактивные HTML-карты, отчеты о покрытии кода и системные журналы."
    )

    p_body(
        "При прохождении тестов формируются следующие артефакты и выходные данные:"
    )

    p_bullet("1. Консольный отчет тестового раннера (Pytest Summary Log):", 
             "Стандартный текстовый лог выполнения тестов, содержащий список выполненных тестовых функций, статус каждого теста (зеленая точка или метка PASSED при успехе; FAILED с детальным assertion failure trace при ошибке), общее число пройденных тестов (181 passed), суммарное время выполнения (например, in 4.25s) и код возврата операционной системы (Exit Code 0 при полном успехе, Exit Code 1 при наличии сбоев).")
    p_bullet("2. Структурированные HTTP-ответы REST API (JSON DTO):", 
             "Каждый интеграционный тест проверяет соответствие статус-кода HTTP (200 OK, 201 Created, 401 Unauthorized, 403 Forbidden, 404 Not Found, 422 Unprocessable Entity, 503 Service Unavailable) и структуры JSON-тела ответа спецификации Pydantic (поля status, data, kpi, recommendation, detail).")
    p_bullet("3. Файлы аналитических отчетов в формате CSV:", 
             "Выходной поток данных (StreamingResponse) при вызове эндпоинтов экспорта. Тесты проверяют наличие заголовка Content-Disposition: attachment, кодировку UTF-8 с BOM-префиксом (байты \\xef\\xbb\\xbf), разделитель-запятую и корректность экранирования строковых значений.")
    p_bullet("4. Сгенерированные геопространственные HTML-карты (Folium / Leaflet):", 
             "Текстовые HTML-документы, генерируемые картографическими сервисами. Тесты проверяют наличие базовых тегов <div class=\"folium-map\">, подключение библиотек Leaflet и корректную сериализацию GeoJSON-координат маркеров филиалов и тепловой карты спроса.")
    p_bullet("5. Метрики покрытия кода тестами (Code Coverage Report):", 
             "Сводный отчет Coverage.py, отражающий процент исполненных строк кода (Statements Coverage > 85%) и веток условий (Branch Coverage) по модулям api, services, repositories и ml.")
    p_bullet("6. Системные логи и сообщения об ошибках:", 
             "Журнал приложения, фиксирующий перехват исключений валидации, предупреждения о перегрузке очередей филиалов и отладочную информацию о ходе загрузки дорожного графа.")

    p_body("В таблице 2.1 представлены характеристики выходных спецификаций:")

    outputs_headers = ["Тип выхода", "Формат / Канал", "Критерий успешности", "Пример выходного значения"]
    outputs_data = [
        ["Лог Pytest", "stdout / stderr", "100% тестов со статусом PASSED, Exit Code = 0", "181 passed in 4.25s"],
        ["Код HTTP", "HTTP Response Status", "Соответствие спецификации REST API v1", "200 OK, 201 Created, 403 Forbidden"],
        ["Тело JSON", "application/json", "Строгая валидация Pydantic DTO", "{\"status\": \"ok\", \"travel_time_min\": 12.4}"],
        ["Экспорт CSV", "text/csv; charset=utf-8", "Наличие UTF-8 BOM, запятая, корректный header", "ORD_ID,Branch,Date,Total\\nORD-1,Centr..."],
        ["Карта Folium", "text/html", "Валидный HTML с внедренным GeoJSON", "<!DOCTYPE html>...folium-map..."],
        ["Coverage", "HTML / XML / text", "Суммарное покрытие кода тестами >= 85%", "TOTAL: 88% statements covered"]
    ]
    p_table(outputs_headers, outputs_data, [1.2, 1.6, 2.0, 1.6])

    # =========================================================================
    # 3. СРЕДА ТЕСТИРОВАНИЯ
    # =========================================================================
    p_h1("3. Среда тестирования")

    p_body(
        "В данном разделе зафиксированы характеристики и конфигурации аппаратного и программного окружения, "
        "в котором выполняется тестирование программного комплекса «NewFastPizza». Процесс обеспечения качества "
        "организован на двух уровнях: на локальных рабочих станциях разработчиков и в автоматизированном "
        "облачном пайплайне непрерывной интеграции (CI/CD GitHub Actions)."
    )

    p_h2("3.1. Локальная тестовая среда разработчика / QA-инженера:")
    p_body(
        "Локальная среда используется для разработки тестов, их отладки и запуска перед отправкой изменений в репозиторий.\n"
        "• Аппаратная платформа: Рабочая станция / ноутбук x86-64 (AMD Ryzen 5 / Intel Core i5/i7, 6+ физических ядер с частотой от 3.0 ГГц);\n"
        "• Оперативная память: 16 ГБ DDR4/DDR5 (не менее 4 ГБ выделяется под выполнение тестового процесса и кэширование графа);\n"
        "• Дисковая подсистема: Твердотельный накопитель NVMe SSD со скоростью чтения от 2000 МБ/с;\n"
        "• Операционная система: Windows 11 Pro (с поддержкой PowerShell 7) или Ubuntu 22.04 LTS Desktop;\n"
        "• Среда интерпретатора: Python 3.12.7 в изолированном виртуальном окружении (.venv);\n"
        "• СУБД: Встроенная SQLite 3 (in-memory) с поддержкой асинхронного драйвера aiosqlite."
    )

    p_h2("3.2. Автоматизированная среда непрерывной интеграции (CI/CD Pipeline):")
    p_body(
        "Каждый коммит и Pull Request в основную ветку main автоматически валидируется в облачной среде GitHub Actions "
        "согласно конфигурационному файлу .github/workflows/ci.yml. Это исключает влияние индивидуальных настроек "
        "компьютеров разработчиков на результаты тестирования."
    )

    p_bullet("• Виртуальная машина раннера (GitHub-hosted Runner):", 
             "ubuntu-latest (актуальный дистрибутив Ubuntu Linux 22.04 LTS x86-64, 2 vCPU, 7 ГБ RAM, 14 ГБ SSD).")
    p_bullet("• Шаг 1: Подготовка окружения (actions/setup-python):", 
             "Развертывание чистой среды Python строго версии 3.12 с кэшированием pip-зависимостей.")
    p_bullet("• Шаг 2: Установка тестовых зависимостей:", 
             "Установка рабочего и тестового окружения из файла backend/requirements-dev.txt (FastAPI, Pytest, HTTPX, Ruff, Scikit-learn, NetworkX, Pandas, NumPy, Folium).")
    p_bullet("• Шаг 3: Статический анализ и линтинг (Ruff Quality Gate):", 
             "Выполнение строгой проверки синтаксиса и стиля кода:\n"
             "ruff check backend\n"
             "ruff format --check backend\n"
             "При обнаружении неиспользуемых импортов, синтаксических ошибок или нарушения форматирования билд прерывается.")
    p_bullet("• Шаг 4: Исполнение полного набора автотестов (Pytest Suite):", 
             "Запуск полного набора модульных и интеграционных тестов:\n"
             "pytest -q --tb=short\n"
             "Все 181 тест должны завершиться со статусом PASSED за время не более 10 секунд.")

    p_callout(
        "Настройка ветки `main` в GitHub репозитории: включено правило Branch Protection Rule. "
        "Слияние любого Pull Request запрещено до тех пор, пока automated check «Ruff and pytest (Python 3.12)» "
        "не завершится с успешным зеленым статусом.",
        title="ПРАВИЛО ЗАЩИТЫ ВЕТКИ MAIN (CI QUALITY GATE):",
        bg_color="ECFDF5",
        border_color="059669"
    )

    # =========================================================================
    # 4. СПЕЦИАЛЬНЫЕ ТРЕБОВАНИЯ
    # =========================================================================
    p_h1("4. Специальные требования")

    p_body(
        "В данном разделе регламентируются специальные требования к организации тестирования, не вошедшие в предыдущие пункты. "
        "В соответствии с разъяснениями преподавателя, идеальным инженерным качеством программной системы является "
        "минимизация или полное отсутствие специфических требований, усложняющих процесс верификации."
    )

    p_body(
        "Для программного комплекса «NewFastPizza» констатируется отсутствие специальных требований к аппаратуре и стендам:\n"
        "1. Отсутствие аппаратных зависимостей: Тестирование не требует подключения физических GPS-трекеров, датчиков или кассовых терминалов. Все пространственные и временные сигналы эмулируются программными генераторами.\n"
        "2. Отсутствие внешних платных сервисов: Тесты не выполняют сетевых обращений к сторонним коммерческим API (Яндекс.Карты, Google Maps, SMS-шлюзы). Это гарантирует 100% автономность выполнения тестового набора даже при полном отсутствии доступа к сети Интернет.\n"
        "3. Изоляция транзакций: Для работы тестов не требуется предварительное развертывание и конфигурирование промышленного сервера СУБД PostgreSQL. Тестовый комплекс самостоятельно поднимает временную базу данных SQLite в памяти, накатывает структуру таблиц и полностью уничтожает ее по завершении прогона."
    )

    # =========================================================================
    # 5. ВЗАИМОЗАВИСИМОСТИ
    # =========================================================================
    p_h1("5. Взаимозависимости")

    p_body(
        "В данном разделе приводятся взаимосвязи между тестами и порядок их выполнения, если он существенен. "
        "В соответствии со строгими правилами объектно-ориентированного тестирования и рекомендациями преподавателя, "
        "все разрабатываемые тесты должны быть взаимно независимыми (изолированными)."
    )

    p_body(
        "В программном комплексе «NewFastPizza» взаимозависимости между тестами отсутствуют. "
        "Каждый тест спроектирован как атомарный и идемпотентный сценарий:\n"
        "1. Изоляция данных на уровне фикстур: Фикстура db_session в conftest.py создает изолированную сессию базы данных для каждого тестового метода. Все модификации данных (INSERT, UPDATE, DELETE) выполняются в рамках локальной транзакции, которая принудительно откатывается (rollback) при выходе из теста.\n"
        "2. Изоляция клиента API: Фикстура client предоставляет чистый экземпляр AsyncClient, гарантирующий отсутствие закэшированных сессионных кук или токенов между последовательными тестами.\n"
        "3. Произвольный порядок запуска: Тесты могут запускаться в любом порядке, выборочно по отдельным модулям (например, pytest backend/tests/test_ml_pipeline.py), а также параллельно в несколько потоков с использованием модуля pytest-xdist (-n auto) без риска взаимного влияния на результаты проверок."
    )

    # =========================================================================
    # 6. ТЕСТЫ
    # =========================================================================
    p_h1("6. Тесты")

    p_body(
        "Раздел «Тесты» является центральной и наиболее объемной частью документа. Раздел строго разделен на две части: "
        "модульное тестирование (Unit Tests) и интеграционное тестирование (Integration Tests). "
        "В соответствии с требованиями преподавателя, раздел содержит подробные описания более 30 модульных тестов, "
        "а также ключевых интеграционных сценариев, подтверждающих корректность сквозного взаимодействия компонентов комплекса. "
        "Для каждого теста по единому академическому стандарту зафиксированы обязательные атрибуты:\n"
        "• Ожидаемое поведение: Входные данные → Выполняемая операция → Ожидаемый результат;\n"
        "• Ожидаемая производительность;\n"
        "• Ожидаемая достоверность;\n"
        "• Замечания."
    )

    # 6.1. Модульное тестирование
    p_h2("6.1. Модульное тестирование (Unit Tests)")
    p_body(
        "Модульные тесты проверяют изолированную работу отдельных классов, функций и методов системы без вовлечения "
        "сетевого взаимодействия и внешних сервисов. Ниже представлено описание 32 ключевых модульных тестов комплекса."
    )

    # --- Подгруппа 1: Репозитории и модели БД ---
    p_h3("Группа 1: Модульные тесты доступа к данным (ORM-модели и репозитории)")

    p_test_case(
        "6.1.1", "Создание пользователя и валидация полей в UserRepository",
        behavior_in="Объект User(username='test_user', email='test@pizza.ru', role='manager', is_active=True)",
        behavior_op="Вызов метода UserRepository.create(user) в тестовой сессии БД",
        behavior_out="Запись успешно сохраняется в таблице users, объекту присваивается целочисленный первичный ключ id > 0, временные метки created_at и updated_at заполнены корректным временем UTC",
        perf="< 5 мс", validity="100%", remarks="Проверяется базовый CRUD-механизм BaseRepository."
    )

    p_test_case(
        "6.1.2", "Ограничение уникальности логина пользователя (Unique Constraint)",
        behavior_in="Попытка создания двух пользователей с идентичным логином username='duplicate_user'",
        behavior_op="Повторный вызов UserRepository.create() со вторым объектом",
        behavior_out="Возбуждается исключение нарушения уникальности IntegrityError, транзакция откатывается, вторая запись не создается",
        perf="< 3 мс", validity="100%", remarks="Проверка ограничения uq_users_username."
    )

    p_test_case(
        "6.1.3", "Проверка CHECK-ограничения роли пользователя (ck_users_role)",
        behavior_in="Объект User с недопустимым строковым значением роли role='super_hacker'",
        behavior_op="Попытка фиксации записи в сессии БД (session.commit())",
        behavior_out="СУБД блокирует операцию с ошибкой IntegrityError (нарушение ограничения допустимых ролей: admin, manager, executive, driver)",
        perf="< 3 мс", validity="100%", remarks="Гарантия строгой ролевой модели на уровне БД."
    )

    p_test_case(
        "6.1.4", "Выборка активных филиалов в BranchRepository",
        behavior_in="В БД создано 3 филиала: 2 с флагом is_active=True и 1 с is_active=False",
        behavior_op="Вызов метода BranchRepository.list_active()",
        behavior_out="Возвращается список строго из 2 активных филиалов, неактивный филиал отфильтрован",
        perf="< 4 мс", validity="100%", remarks="Исключение закрытых филиалов из логистики."
    )

    p_test_case(
        "6.1.5", "Ограничение неотрицательности очереди филиала (ck_branches_current_queue)",
        behavior_in="Объект Branch с отрицательным значением очереди current_queue = -5",
        behavior_op="Попытка сохранения записи в базу данных",
        behavior_out="СУБД выбрасывает IntegrityError, значение очереди не может быть меньше 0",
        perf="< 3 мс", validity="100%", remarks="Защита от сбоя счетчика заказов на кухне."
    )

    p_test_case(
        "6.1.6", "Привязка курьера к филиалу и пользователю в DriverRepository",
        behavior_in="Созданный пользователь User(id=10, role='driver') и филиал Branch(id=2)",
        behavior_op="Создание Driver(user_id=10, branch_id=2, full_name='Иванов И.И.')",
        behavior_out="Курьер успешно сохраняется, внешние ключи user_id и branch_id связывают сущности, метод get_by_user_id(10) возвращает созданный объект",
        perf="< 5 мс", validity="100%", remarks="Проверка целостности связей 1-к-1 и 1-ко-многим."
    )

    p_test_case(
        "6.1.7", "Фильтрация выборки заказов по временному диапазону в OrderRepository",
        behavior_in="Массив из 10 заказов с датами от 01.11.2026 до 10.11.2026; запрос с фильтром: start=03.11, end=06.11",
        behavior_op="Вызов OrderRepository.get_orders_in_range(start_date, end_date)",
        behavior_out="Возвращаются только заказы, попадающие в полуинтервал [03.11, 06.11), заказы вне диапазона исключены",
        perf="< 8 мс", validity="100%", remarks="Используются составные индексы ix_orders_order_time."
    )

    p_test_case(
        "6.1.8", "Пагинация и удаление сущностей в BaseRepository",
        behavior_in="Таблица с 25 записями; запрос страницы с параметрами offset=10, limit=5",
        behavior_op="Вызов list_paged(offset=10, limit=5) с последующим delete(id=15)",
        behavior_out="Метод возвращает ровно 5 записей (с 11 по 15); последующее удаление возвращает True, и запись с id=15 исчезает из БД",
        perf="< 6 мс", validity="100%", remarks="Проверка базовой функциональности пагинации."
    )

    # --- Подгруппа 2: Безопасность и криптография ---
    p_h3("Группа 2: Модульные тесты модуля безопасности и аутентификации")

    p_test_case(
        "6.1.9", "Хэширование пароля алгоритмом Bcrypt",
        behavior_in="Строка исходного пароля в открытом виде: 'SecretPassword123!'",
        behavior_op="Вызов функции get_password_hash(password)",
        behavior_out="Возвращается строка хэша длиной 60 символов, начинающаяся с идентификатора алгоритма '$2b$', открытый пароль в строке отсутствует",
        perf="< 80 мс", validity="100%", remarks="Трудоемкость Bcrypt настроена на 12 раундов."
    )

    p_test_case(
        "6.1.10", "Верификация корректного и некорректного пароля",
        behavior_in="Хэш пароля 'SecretPassword123!' и две строки: правильная и ошибочная 'WrongPass'",
        behavior_op="Вызовы verify_password('SecretPassword123!', hash) и verify_password('WrongPass', hash)",
        behavior_out="Первый вызов возвращает True, второй вызов возвращает False",
        perf="< 80 мс", validity="100%", remarks="Защита от атак по времени (constant time comparison)."
    )

    p_test_case(
        "6.1.11", "Генерация JWT-токена доступа (Access Token)",
        behavior_in="Словарь данных: {'sub': 'ivan_mgr', 'role': 'manager', 'user_id': 5}",
        behavior_op="Вызов create_access_token(data, expires_delta=timedelta(minutes=60))",
        behavior_out="Возвращается валидная закодированная JWT-строка из трех частей (header.payload.signature), заголовок содержит алгоритм HS256",
        perf="< 1 мс", validity="100%", remarks="Формирование токена по стандарту RFC 7519."
    )

    p_test_case(
        "6.1.12", "Проверка истечения срока действия токена (Token Expiration)",
        behavior_in="JWT-токен с отрицательным сроком жизни expires_delta = timedelta(seconds=-10)",
        behavior_op="Попытка декодирования токена через decode_access_token(token)",
        behavior_out="Возбуждается исключение jwt.ExpiredSignatureError, токен признается недействительным",
        perf="< 1 мс", validity="100%", remarks="Гарантия невозможности использования просроченных токенов."
    )

    p_test_case(
        "6.1.13", "Отклонение токена с неверной цифровой подписью",
        behavior_in="Валидный JWT-токен, подписанный сторонним секретным ключом 'fake_secret_key'",
        behavior_op="Попытка верификации системным сервисом security с ключом JWT_SECRET_KEY",
        behavior_out="Возбуждается исключение jwt.InvalidSignatureError, доступ блокируется",
        perf="< 1 мс", validity="100%", remarks="Защита от подделки токенов злоумышленниками."
    )

    p_test_case(
        "6.1.14", "Проверка ролевого фильтра доступа (Role Hierarchy)",
        behavior_in="Пользователь с ролью 'driver' и функция, требующая роли ['admin', 'manager']",
        behavior_op="Вызов защитного декоратора require_roles(user, allowed_roles)",
        behavior_out="Возбуждается HTTPException со статусом 403 Forbidden, выполнение функции прерывается",
        perf="< 0.5 мс", validity="100%", remarks="Проверка изоляции функций RBAC."
    )

    # --- Подгруппа 3: Статистика и KPI ---
    p_h3("Группа 3: Модульные тесты сервиса статистики и агрегации показателей")

    p_test_case(
        "6.1.15", "Точный расчет KPI на фиксированном наборе данных в StatsService",
        behavior_in="DataFrame из 5 заказов: 4 доставлены (время: 20, 30, 40, 50 мин), 1 отменен",
        behavior_op="Вызов StatsService.calculate_summary_kpi(df)",
        behavior_out="Возвращается объект KPI: total_orders=5, delivered_orders=4, canceled_orders=1, avg_delivery_time=35.0 мин, cancel_rate=20.0%",
        perf="< 3 мс", validity="100%", remarks="Векторизованные математические вычисления в Pandas."
    )

    p_test_case(
        "6.1.16", "Расчет KPI на пустом наборе данных (Zero Division Guard)",
        behavior_in="Пустой DataFrame (0 строк) с сохранением структуры колонок",
        behavior_op="Вызов StatsService.calculate_summary_kpi(empty_df)",
        behavior_out="Возвращается объект с нулевыми значениями: total_orders=0, avg_delivery_time=0.0, ошибки деления на ноль не происходит",
        perf="< 1 мс", validity="100%", remarks="Защита от аварийного падения при отсутствии заказов."
    )

    p_test_case(
        "6.1.17", "Расчет KPI при 100% отмененных заказов",
        behavior_in="DataFrame из 3 заказов, все имеют статус status='Canceled'",
        behavior_op="Вызов StatsService.calculate_summary_kpi(canceled_df)",
        behavior_out="total_orders=3, delivered=0, canceled=3, cancel_rate=100.0%, avg_delivery_time=0.0",
        perf="< 2 мс", validity="100%", remarks="Проверка корректности фильтрации по статусам."
    )

    p_test_case(
        "6.1.18", "Расчет загрузки филиалов (Branch Load Aggregation)",
        behavior_in="DataFrame с 10 заказами, распределенными по трем филиалам (Branch 1: 5, Branch 2: 3, Branch 3: 2)",
        behavior_op="Вызов StatsService.calculate_branch_load(df)",
        behavior_out="Возвращается список из 3 элементов, значения count совпадают с фактическими, доли (percentage) составляют 50%, 30% и 20%",
        perf="< 3 мс", validity="100%", remarks="Проверка группировки groupby('branch_id')."
    )

    p_test_case(
        "6.1.19", "Построение таймсерии с заполнением пропусков (Fill Gaps)",
        behavior_in="Заказы присутствуют только за 01.11 и 03.11; интервал анализа с 01.11 по 03.11 с шагом 1 день",
        behavior_op="Вызов StatsService.build_timeseries(df, granularity='day')",
        behavior_out="В результирующем массиве присутствуют 3 точки: 01.11 (count>0), 02.11 (count=0, искусственно заполненный пропуск), 03.11 (count>0)",
        perf="< 5 мс", validity="100%", remarks="Непрерывность отображения графиков на фронтенде."
    )

    p_test_case(
        "6.1.20", "Агрегация заказов по часам суток (Hourly Distribution)",
        behavior_in="Массив заказов со временем оформления от 10:00 до 23:00",
        behavior_op="Вызов StatsService.aggregate_by_hour(df)",
        behavior_out="Возвращается словарь из 24 элементов (часы 0..23), где каждому часу сопоставлено точное количество поступивших заказов",
        perf="< 3 мс", validity="100%", remarks="Анализ пиковых часов нагрузки на кухню."
    )

    # --- Подгруппа 4: Отчеты и CSV ---
    p_h3("Группа 4: Модульные тесты генератора отчетов и экспорта CSV")

    p_test_case(
        "6.1.21", "Валидация JSON-схемы шаблона отчета в ReportConfig",
        behavior_in="Конфигурация {'columns': ['order_id', 'delivery_time'], 'date_range': {'days': 7}}",
        behavior_op="Валидация схемы через Pydantic-модель ReportConfig",
        behavior_out="Схема успешно валидируется, создается объект конфигурации с типизированными полями",
        perf="< 1 мс", validity="100%", remarks="Проверка формата пользовательских настроек."
    )

    p_test_case(
        "6.1.22", "Отклонение некорректной конфигурации шаблона отчета",
        behavior_in="Недопустимый JSON с отрицательным диапазоном дней: {'date_range': {'days': -10}}",
        behavior_op="Попытка парсинга через ReportConfig",
        behavior_out="Возбуждается pydantic.ValidationError с указанием на недопустимое значение поля days",
        perf="< 1 мс", validity="100%", remarks="Защита от некорректных параметров фильтрации."
    )

    p_test_case(
        "6.1.23", "Генерация CSV-строки с разделителем-запятой",
        behavior_in="Таблица из 2 строк данных: [['ORD-01', 'Центральный', 25], ['ORD-02', 'Сормово', 40]]",
        behavior_op="Вызов функции format_csv_stream(headers, rows)",
        behavior_out="Генерируется текстовый поток, строки разделены символом '\\r\\n', колонки разделены запятыми, кириллические названия филиалов не повреждены",
        perf="< 2 мс", validity="100%", remarks="Соответствие стандарту RFC 4180."
    )

    p_test_case(
        "6.1.24", "Экранирование спецсимволов и кавычек в CSV-экспорте",
        behavior_in="Строка данных, содержащая запятую и кавычки: 'Пиццерия \"Люкс\", филиал №1'",
        behavior_op="Вызов хелпера экранирования csv_escape()",
        behavior_out="Поле оборачивается в двойные кавычки, внутренние кавычки дублируются: '\"Пиццерия \"\"Люкс\"\", филиал №1\"'",
        perf="< 0.5 мс", validity="100%", remarks="Предотвращение поломки структуры CSV при импорте в Excel."
    )

    # --- Подгруппа 5: Графы и алгоритм Дейкстры ---
    p_h3("Группа 5: Модульные тесты алгоритмов на графах и маршрутизации")

    p_test_case(
        "6.1.25", "Поиск кратчайшего пути Дейкстры на эталонном графе",
        behavior_in="Синтетический граф дорог (7 узлов); начальный узел 1, целевой узел 5",
        behavior_op="Вызов RoutingService.calculate_shortest_path(graph, source=1, target=5, weight='travel_time')",
        behavior_out="Возвращается список узлов оптимального пути [1, 4, 5] с суммарным временем в пути 400.0 секунд (кратчайший путь)",
        perf="< 2 мс", validity="100%", remarks="Математическая проверка реализации алгоритма Дейкстры."
    )

    p_test_case(
        "6.1.26", "Учет коэффициентов пробок при поиске кратчайшего пути",
        behavior_in="На ребре (1, 4) увеличен вес пробки в 10 раз (multiplier=10.0)",
        behavior_op="Повторный запуск алгоритма Дейкстры от узла 1 до узла 5",
        behavior_out="Маршрут перестраивается в обход затора через узлы [1, 2, 5], подтверждая адаптивность алгоритма к пробкам",
        perf="< 2 мс", validity="100%", remarks="Динамический учет дорожной ситуации."
    )

    p_test_case(
        "6.1.27", "Многокритериальный скоринг филиалов с учетом очередей",
        behavior_in="Два филиала: Филиал А (время пути 10 мин, очередь 10 заказов) и Филиал Б (время пути 15 мин, очередь 0 заказов)",
        behavior_op="Вызов RoutingAdapter.evaluate_branch_scores(branches_eval_data)",
        behavior_out="Филиал Б получает более высокий приоритет (меньший суммарный штраф), так как отсутствие очереди компенсирует лишние 5 минут езды",
        perf="< 1 мс", validity="100%", remarks="Минимизация времени простоя курьера."
    )

    p_test_case(
        "6.1.28", "Жизненный цикл графа дорог в оперативной памяти (GraphState)",
        behavior_in="Файл графа test_graph.pkl; обращение к singleton-объекту GraphState",
        behavior_op="Последовательные вызовы: is_loaded(), load(), get_graph(), unload()",
        behavior_out="До загрузки is_loaded()=False; после загрузки возвращается граф с 7 узлами; после выгрузки память освобождается",
        perf="< 5 мс", validity="100%", remarks="Управление ресурсами RAM сервера."
    )

    # --- Подгруппа 6: Машинное обучение K-Means ---
    p_h3("Группа 6: Модульные тесты машинного обучения и синтеза данных")

    p_test_case(
        "6.1.29", "Сходимость алгоритма кластеризации K-Means на геоданных",
        behavior_in="Массив из 100 географических координат доставок, сгруппированных вокруг 3 выраженных центров спроса",
        behavior_op="Вызов KMeansAnalyzer.fit_predict(coords, n_clusters=3)",
        behavior_out="Алгоритм успешно сходится за < 20 итераций, вычисляет координаты 3 центроидов, близкие к эталонным центрам генерации",
        perf="< 25 мс", validity="100%", remarks="Кластеризация Scikit-learn."
    )

    p_test_case(
        "6.1.30", "Обработка недостаточного объема выборки в K-Means (N < K)",
        behavior_in="Массив из 3 координат доставок при запросе K=7 кластеров",
        behavior_op="Попытка вызова KMeansAnalyzer.analyze(coords, n_clusters=7)",
        behavior_out="Возбуждается доменное исключение InsufficientDataError с понятным сообщением «Количество точек меньше числа кластеров»",
        perf="< 1 мс", validity="100%", remarks="Предотвращение краша математической модели."
    )

    p_test_case(
        "6.1.31", "Генерация синтетических координат заказов в полигоне города",
        behavior_in="Запрос генерации 50 заказов для Нижнего Новгорода",
        behavior_op="Вызов генератора synthetic_order_generator(count=50)",
        behavior_out="Все 50 сгенерированных точек имеют широту lat в диапазоне [56.20, 56.40] и долготу lon в [43.80, 44.10], координаты не выходят за границы полигона",
        perf="< 10 мс", validity="100%", remarks="Реалистичность синтетических данных."
    )

    p_test_case(
        "6.1.32", "Идемпотентность скрипта первичного наполнения БД (Seed Idempotency)",
        behavior_in="Пустая БД; двукратный последовательный запуск скрипта seed.py",
        behavior_op="Первый запуск: заполнение таблиц; второй запуск: повторная попытка наполнения",
        behavior_out="Первый запуск создает 7 филиалов и администратора; второй запуск определяет наличие данных и не создает дубликатов (число записей неизменно)",
        perf="< 150 мс", validity="100%", remarks="Безопасность повторного запуска сида."
    )

    # 6.2. Интеграционное тестирование
    p_h2("6.2. Интеграционное тестирование (Integration Tests)")
    p_body(
        "Интеграционные тесты проверяют корректность совместного функционирования компонентов системы: взаимодействие "
        "REST API, базы данных, сервисного слоя, алгоритмов машинного обучения и подсистемы безопасности при выполнении "
        "сквозных пользовательских сценариев. Для каждого теста зафиксированы жесткие требования к производительности "
        "(время отклика) и критерии достоверности."
    )

    p_test_case(
        "6.2.1", "Сквозная регистрация нового пользователя в системе",
        behavior_in="HTTP-запрос POST /api/v1/auth/register с JSON-телом: {'username': 'new_manager', 'email': 'mgr@pizza.ru', 'password': 'SecurePassword123', 'role': 'manager'}",
        behavior_op="API валидирует тело через Pydantic DTO → AuthService хэширует пароль Bcrypt → UserRepository сохраняет запись в БД",
        behavior_out="HTTP статус 201 Created; тело ответа содержит {'id': 1, 'username': 'new_manager', 'role': 'manager', 'is_active': true}; хэш пароля в ответе отсутствует; в БД появилась новая запись",
        perf="< 120 мс (с учетом раундов Bcrypt)", validity="100%", remarks="Сквозной сценарий создания пользователя."
    )

    p_test_case(
        "6.2.2", "Аутентификация пользователя и выпуск Bearer JWT токена",
        behavior_in="HTTP-запрос POST /api/v1/auth/jwt/login с form-data: username='new_manager', password='SecurePassword123'",
        behavior_op="AuthRouter передает данные в AuthService → извлечение пользователя из БД → верификация пароля → выпуск JWT",
        behavior_out="HTTP статус 200 OK; тело содержит {'access_token': '<jwt_string>', 'token_type': 'bearer'}; токен успешно валидируется секретным ключом",
        perf="< 100 мс", validity="100%", remarks="Проверка протокола OAuth2 Password Flow."
    )

    p_test_case(
        "6.2.3", "Ролевая изоляция эндпоинтов (RBAC Access Control)",
        behavior_in="HTTP-запрос GET /api/v1/stats/summary с заголовком Authorization: Bearer <driver_token> (токен с ролью 'driver')",
        behavior_op="FastAPI Dependency require_roles(['admin', 'manager', 'executive']) декодирует токен и проверяет права роли",
        behavior_out="HTTP статус 403 Forbidden; тело содержит {'detail': 'Недостаточно прав для выполнения операции'}; запрос блокируется до вызова сервиса",
        perf="< 15 мс", validity="100%", remarks="Защита аналитических данных от несанкционированного доступа курьеров."
    )

    p_test_case(
        "6.2.4", "Сквозной расчет аналитических показателей KPI дашборда",
        behavior_in="HTTP-запрос GET /api/v1/stats/summary?start_date=2026-11-01&end_date=2026-11-15 от имени менеджера; в БД загружено 100 заказов",
        behavior_op="Контроллер проверяет права → OrderRepository выполняет SQL-запрос с фильтрацией по индексам → StatsService рассчитывает KPI в Pandas",
        behavior_out="HTTP статус 200 OK; тело ответа содержит валидный JSON со структурой: total_orders, delivered_orders, avg_delivery_time_min, cancel_rate; все математические соотношения соблюдены",
        perf="< 80 мс", validity="100%", remarks="Сквозной интеграционный сценарий экрана Dashboard."
    )

    p_test_case(
        "6.2.5", "Анализ распределения нагрузки по филиалам сети",
        behavior_in="HTTP-запрос GET /api/v1/stats/branch-load от авторизованного менеджера",
        behavior_op="Выборка активных филиалов из BranchRepository → связка с заказами из OrderRepository → расчет долей в StatsService",
        behavior_out="HTTP статус 200 OK; список всех 7 филиалов сети с указанием количества заказов, суммарной выручки и процента от общего потока; сумма долей равна 100%",
        perf="< 60 мс", validity="100%", remarks="Проверка корректности реляционного объединения (JOIN)."
    )

    p_test_case(
        "6.2.6", "Формирование временных рядов динамики заказов (Timeseries)",
        behavior_in="HTTP-запрос GET /api/v1/stats/timeseries?granularity=day&start_date=2026-11-01&end_date=2026-11-07",
        behavior_op="OrderRepository извлекает агрегаты по дням → StatsService заполняет дни без заказов нулевыми точками",
        behavior_out="HTTP статус 200 OK; массив ровно из 7 точек дат, упорядоченных по возрастанию; отсутствие пропусков в датах",
        perf="< 45 мс", validity="100%", remarks="Интеграция с библиотекой графиков Chart.js."
    )

    p_test_case(
        "6.2.7", "Создание, сохранение и извлечение шаблона отчета",
        behavior_in="HTTP-запрос POST /api/v1/reports/templates с конфигурацией колонок и фильтров от менеджера",
        behavior_op="Валидация JSON через ReportTemplateCreate → ReportTemplateRepository сохраняет шаблон в БД с привязкой owner_id",
        behavior_out="HTTP статус 201 Created; последующий запрос GET /api/v1/reports/templates возвращает сохраненный шаблон в списке персональных отчетов пользователя",
        perf="< 30 мс", validity="100%", remarks="Проверка сохранения сложных JSON-структур в реляционной таблице."
    )

    p_test_case(
        "6.2.8", "Сквозной потоковый экспорт аналитического отчета в формате CSV",
        behavior_in="HTTP-запрос POST /api/v1/reports/export/csv с телом параметров фильтрации",
        behavior_op="ReportService формирует DataFrame → преобразует в текстовый поток CSV с UTF-8 BOM → StreamingResponse отдает поток",
        behavior_out="HTTP статус 200 OK; заголовок Content-Type: text/csv; charset=utf-8; первые 3 байта равны \\xef\\xbb\\xbf; первая строка содержит корректные заголовки колонок",
        perf="< 150 мс при выборке до 5000 строк", validity="100%", remarks="Проверка потоковой выгрузки без переполнения памяти."
    )

    p_test_case(
        "6.2.9", "Сквозной расчет рекомендации оптимального филиала курьеру",
        behavior_in="HTTP-запрос POST /api/v1/recommendations/driver-branch с координатами курьера {'lat': 56.3269, 'lon': 44.0075}",
        behavior_op="API валидирует координаты → GraphService находит ближайший узел дорожного графа → RoutingService запускает Дейкстру до всех филиалов → учет очередей",
        behavior_out="HTTP статус 200 OK; ответ содержит: recommended_branch_id, branch_name, travel_time_minutes, distance_km, polyline_nodes; выбран действительно оптимальный филиал",
        perf="< 90 мс на полном графе Нижнего Новгорода", validity="98% (с учетом динамики пробок)", remarks="Ключевой логистический сценарий системы."
    )

    p_test_case(
        "6.2.10", "Пространственный ML-анализ зон спроса K-Means",
        behavior_in="HTTP-запрос GET /api/v1/recommendations/branch-clusters?n_clusters=7&period_days=30 от менеджера",
        behavior_op="DemandAdapter извлекает массивы доставок из БД → обучает K-Means Scikit-learn → сопоставляет центроиды с филиалами",
        behavior_out="HTTP статус 200 OK; тело содержит 7 кластеров с координатами центроидов, радиусами, плотностью и рекомендациями (OPEN_NEW_BRANCH / OPTIMIZE_BRANCH / OPTIMAL)",
        perf="< 250 мс", validity="95%", remarks="Сквозная интеграция модуля Data Science с веб-API."
    )

    p_test_case(
        "6.2.11", "Калибровка весов дорожного графа администратором",
        behavior_in="HTTP-запрос POST /api/v1/admin/graph/weights от администратора с прикрепленным файлом weights.csv (коэффициенты пробок)",
        behavior_op="GraphAdminService парсит CSV → обновляет веса ребер графа в оперативной памяти → инвалидирует кэш маршрутов",
        behavior_out="HTTP статус 200 OK; ответ содержит {'updated_edges': 15, 'status': 'applied'}; последующий вызов маршрутизации курьера строит путь с учетом новых пробок",
        perf="< 180 мс", validity="100%", remarks="Динамическое управление графом без перезагрузки сервера."
    )

    p_test_case(
        "6.2.12", "Комплексная проверка готовности сервиса (Health Check Endpoint)",
        behavior_in="HTTP-запрос GET /api/v1/health/ready без авторизации (публичный служебный мониторинг)",
        behavior_op="HealthRouter выполняет проверочный запрос к БД (SELECT 1) и проверяет флаг graph_state.is_loaded()",
        behavior_out="HTTP статус 200 OK; тело ответа: {'status': 'ok', 'database': 'connected', 'graph': 'ready'}; при имитации отказа БД возвращается статус 503 Service Unavailable",
        perf="< 10 мс", validity="100%", remarks="Используется для Docker healthcheck и CI/CD smoke-test.sh."
    )

    # Save final document
    out_dir = "docs"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "4. Спецификации и сценарии тестов.docx")
    doc.save(out_path)
    print(f"Document 4 successfully created and saved to: {out_path}")
    return out_path

if __name__ == '__main__':
    create_document()


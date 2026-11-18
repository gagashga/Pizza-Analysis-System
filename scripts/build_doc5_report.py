# -*- coding: utf-8 -*-
"""
Builder for Document 5: Testing and Bug Reports (Тестирование и отчеты об ошибках)
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

    def p_callout(text, title="СТАБИЛИЗАЦИЯ И РЕЛИЗНЫЙ СТАТУС БИЛДА:", bg_color="FEF3C7", border_color="D97706"):
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

    def p_image(img_path, caption, width_in=4.5):
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
        p_cap.paragraph_format.space_after = Pt(10)
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
    r = p_title.add_run("ТЕСТИРОВАНИЕ И ОТЧЕТЫ ОБ ОШИБКАХ\n")
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
        ("1. Введение", "2"),
        ("2. Методы и инструменты тестирования", "3"),
        ("    2.1. Обзор методов тестирования", "3"),
        ("    2.2. Обзор средств тестирования", "5"),
        ("3. Отчет о тестах", "7"),
        ("    3.1. Результаты тестов – разница от прошлого отчета", "7"),
        ("    3.2. Отклонения в выполнении плана тестирования", "9"),
        ("    3.3. Оставшиеся тесты и исправления в плане тестирования", "10"),
        ("4. Зоны тестирования", "11"),
        ("    4.1. Зона 1: Модуль аутентификации и ролевого доступа (Auth & RBAC)", "11"),
        ("    4.2. Зона 2: Модуль агрегации KPI и аналитики заказов (Stats & Analytics)", "13"),
        ("    4.3. Зона 3: Модуль маршрутизации курьеров на графе дорог (Routing & Graphs)", "15"),
        ("    4.4. Зона 4: Модуль пространственного машинного обучения K-Means", "17"),
        ("    4.5. Зона 5: Конструктор отчетов и потоковый экспорт CSV", "19"),
        ("    4.6. Зона 6: Клиентский интерфейс веб-дашборда и карты (Frontend SPA)", "21"),
        ("5. Отчет об ошибках", "23"),
        ("    5.1. Статус билда", "23"),
        ("    5.2. Известные ошибки", "24"),
        ("    5.3. Анализ сходимости ошибок", "26"),
        ("6. Обновленный план тестирования", "30"),
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
    # 1. ВВЕДЕНИЕ
    # =========================================================================
    p_h1("1. Введение")

    p_body(
        "Данный документ представляет собой комплексный отчет о тестировании и анализе дефектов программного комплекса "
        "«NewFastPizza» (система сбора метрик заказов, генерации аналитической отчетности, маршрутизации курьеров "
        "и оптимизации филиальной сети). Документ обобщает результаты верификации решения на всех фазах жизненного цикла разработки: "
        "Разработки (Developing), Тестирования (Testing) и Стабилизации (Stabilizing) в соответствии с методологией MSF "
        "(Microsoft Solutions Framework)."
    )

    p_body(
        "Важность настоящего документа и высокая частота обращений к нему проектной группы обусловлены тем, что качество "
        "разрабатываемого решения, стабильность серверного ядра и математическая корректность выдаваемых рекомендаций "
        "непосредственно определяются результатами испытаний. Документ служит объективным индикатором готовности системы "
        "к переходу в фазу развертывания (Deploying) и проведению приемо-сдаточных демонстраций заказчику."
    )

    p_body(
        "Структура документа включает следующие основные разделы:\n"
        "• Раздел 2 «Методы и инструменты тестирования»: обзор примененных методов обеспечения качества (модульное, "
        "интеграционное, регрессионное, нагрузочное, тестирование безопасности) и инструментальных средств (Pytest, HTTPX, SQLite, Ruff, CI/CD);\n"
        "• Раздел 3 «Отчет о тестах»: количественная динамика прохождения тестов по итерациям, анализ отклонений от первоначального плана и обоснование решений команды;\n"
        "• Раздел 4 «Зоны тестирования»: детальный анализ результатов испытаний по 6 выделенным функциональным зонам системы "
        "(цели, критерии, результаты, рекомендации);\n"
        "• Раздел 5 «Отчет об ошибках»: оценка статуса текущего билда, реестр выявленных и устраненных дефектов, "
        "математический анализ динамики ошибок с фиксацией точек конвергенции (Bug Convergence) и достижения нуля (Zero-Bug Bounce);\n"
        "• Раздел 6 «Обновленный план тестирования»: финальная корректировка плана верификации перед релизом продукта."
    )

    # =========================================================================
    # 2. МЕТОДЫ И ИНСТРУМЕНТЫ ТЕСТИРОВАНИЯ
    # =========================================================================
    p_h1("2. Методы и инструменты тестирования")

    p_body(
        "Раздел содержит детальную информацию о методологии и инструментальном арсенале тестирования, сопоставляя "
        "фактически выполненные мероприятия с плановыми показателями. Приводятся сведения об используемых подходах, "
        "обоснование выбора средств автоматизации и анализ дополнительных мероприятий по обеспечению надежности решения."
    )

    # 2.1. Обзор методов тестирования
    p_h2("2.1. Обзор методов тестирования")
    p_body(
        "В проекте реализован многоуровневый комплекс методов тестирования, покрывающий как изолированные алгоритмические модули, "
        "так и сквозные бизнес-сценарии использования системы:"
    )

    p_bullet("1. Модульное тестирование (Unit Testing):", 
             "Метод верификации наименьших изолированных компонентов системы: функций хэширования паролей, генераторов токенов JWT, методов репозиториев (UserRepository, BranchRepository, OrderRepository), векторизованных расчетов KPI в Pandas и базовых алгоритмов на графах. Позволяет мгновенно локализовать дефекты на уровне конкретных строк кода.")
    p_bullet("2. Интеграционное тестирование (Integration Testing):", 
             "Проверка корректности совместной работы слоев архитектуры (REST API контроллеры ↔ Сервисный слой ↔ Слой репозиториев ↔ База данных). Выполняется через асинхронный клиент HTTPX ASGI Transport без накладных расходов физической сети.")
    p_bullet("3. Регрессионное тестирование (Regression Testing):", 
             "Автоматизированный перезапуск полного набора из 181 теста при каждом коммите и создании Pull Request в ветку main. Гарантирует, что добавление нового функционала или исправление багов не нарушает ранее стабилизированные компоненты.")
    p_bullet("4. Тестирование безопасности и разграничения доступа (Security Testing):", 
             "Специализированная проверка ролевой модели RBAC (Role-Based Access Control). Эмулируются попытки несанкционированного доступа курьеров к сводкам KPI и генератору отчетов (проверка возврата HTTP 403 Forbidden), проверка устойчивости к поддельным и просроченным JWT-токенам.")
    p_bullet("5. Тестирование граничных условий (Boundary & Edge Case Testing):", 
             "Проверка устойчивости системы в нештатных ситуациях: нулевое количество заказов в выборке (защита от ZeroDivisionError), 100% отмененных заказов, кластеризация K-Means при числе точек меньше заданного числа кластеров (N < K), изолированные вершины дорожного графа.")
    p_bullet("6. Нагрузочное и стресс-тестирование (Performance Testing):", 
             "Оценка времени отклика ключевых алгоритмических эндпоинтов на реальном дорожном графе города Нижнего Новгорода (тысячи перекрестков и дорожных сегментов) и массивах данных до 10 000 заказов. Время поиска маршрута Дейкстры не превышает 100 мс.")
    p_bullet("7. Сквозное приемочное тестирование интерфейса (UI/E2E Testing):", 
             "Ручная проверка корректности отображения аналитического дашборда на Vue 3, плавности отрисовки графиков Chart.js и корректности рендеринга слоев карт Folium/Leaflet в браузерах Chrome, Firefox, Safari и Edge.")

    # 2.2. Обзор средств тестирования
    p_h2("2.2. Обзор средств тестирования")
    p_body(
        "Все примененные средства тестирования относятся к современному стеку Open Source технологий, "
        "обеспечивают 100% повторяемость результатов и полностью интегрированы в процесс разработки:"
    )

    tools_headers = ["Инструмент / Средство", "Тип средства", "Применяемый метод", "Роль в проекте «NewFastPizza»"]
    tools_data = [
        ["Pytest 8.3+", "Тестовый раннер / Фреймворк", "Unit & Integration", "Центральный инструмент: обнаружение, фикстуры, параметризация, запуск 181 теста."],
        ["HTTPX 0.28+", "ASGI HTTP Client", "Интеграционное тестирование", "Эмуляция HTTP-запросов к FastAPI в памяти без поднятия сокета."],
        ["SQLite In-Memory", "Встроенная реляционная СУБД", "Unit & Integration", "Изолированное хранилище данных, создаваемое на лету с авто-откатом транзакций."],
        ["Ruff 0.8+", "Статический анализатор / Линтер", "Статическое тестирование", "Автоматический контроль соответствия кода стандартам PEP 8 и проверка типизации."],
        ["GitHub Actions", "CI/CD Окружение", "Регрессионное тестирование", "Автоматический запуск проверок качества и тестов при каждом Pull Request."],
        ["Folium & Leaflet", "Геоинформационные библиотеки", "Визуальное тестирование карт", "Генерация интерактивных карт маршрутов и тепловых карт спроса."],
        ["smoke-test.sh", "Bash-скрипт мониторинга", "Smoke Testing / Деплой", "Проверка готовности сервиса (/api/v1/health/ready) после развертывания на VPS."],
        ["Swagger UI (/docs)", "Интерактивная консоль API", "Исследовательское тестирование", "Ручная отладка и проверка форматов входных/выходных схем Pydantic DTO."]
    ]
    p_table(tools_headers, tools_data, [1.5, 1.4, 1.4, 2.1])

    # =========================================================================
    # 3. ОТЧЕТ О ТЕСТАХ
    # =========================================================================
    p_h1("3. Отчет о тестах")

    p_body(
        "Раздел обобщает количественные и качественные результаты проведенного тестирования, динамику изменения "
        "показателей успешности отчетов по мере развития проекта, анализ отклонений от первоначального плана "
        "и перечень скорректированных задач."
    )

    # 3.1. Результаты тестов – разница от прошлого отчета
    p_h2("3.1. Результаты тестов – разница от прошлого отчета")
    p_body(
        "Тестирование программного комплекса велось итеративно в течение ноября–декабря 2026 года. "
        "В таблице 3.1 представлена сравнительная динамика ключевых показателей тестирования по трем контрольным срезам:\n"
        "• Отчет 1 (Конец ноября 2026 г., Спринт 2): завершение разработки базового бэкенда, моделей БД и аутентификации;\n"
        "• Отчет 2 (Середина декабря 2026 г., Спринт 3): реализация ML-алгоритмов, сервиса отчетов и интеграции с фронтендом;\n"
        "• Отчет 3 (Конец декабря 2026 г., Спринт 4 — Финальный): полная стабилизация релиз-кандидата Build v1.0.0-rc3."
    )

    progress_headers = ["Контрольный срез / Этап", "Всего тестов", "Успешно (Passed)", "Сбои (Failed)", "Покрытие кода", "Статус релиза"]
    progress_data = [
        ["Отчет 1 (Спринт 2, 23.11.2026)", "45", "38 (84.4%)", "7 (15.6%)", "62.4%", "Alpha: активная разработка ядра"],
        ["Отчет 2 (Спринт 3, 11.12.2026)", "120", "112 (93.3%)", "8 (6.7%)", "78.1%", "Beta: интеграция ML и дашборда"],
        ["Отчет 3 (Спринт 4, 25.12.2026)", "181", "181 (100.0%)", "0 (0.0%)", "88.6%", "Release Candidate: СТАБИЛЕН"]
    ]
    p_table(progress_headers, progress_data, [1.8, 0.9, 1.2, 1.0, 1.0, 1.5])

    p_body(
        "Анализ разницы между прошлым и текущим отчетами:\n"
        "1. Количественный прирост: Набор тестов увеличен со 120 до 181 теста (+61 тест). Прирост обеспечен глубоким покрытием граничных условий K-Means кластеризации, динамической калибровки весов дорожного графа и стресс-тестов потокового экспорта CSV.\n"
        "2. Устранение всех сбоев: Все 8 дефектов, зафиксированных в отчете 2, были полностью локализованы и закрыты. В текущей версии 0 упавших тестов (100% Pass Rate).\n"
        "3. Рост покрытия кода: Суммарное покрытие кодовой базы тестами выросло с 78.1% до 88.6%, при этом критические модули безопасности (security) и маршрутизации курьеров покрыты на 96%."
    )

    # 3.2. Отклонения в выполнении плана тестирования
    p_h2("3.2. Отклонения в выполнении плана тестирования")
    p_body(
        "В ходе практической реализации тестового комплекса проектная группа внесла два обоснованных отклонения "
        "от исходного плана тестирования, позволивших существенно повысить надежность и скорость сборки:"
    )

    p_bullet("• Отклонение 1: Отказ от интеграционного тестирования против внешнего онлайн-сервера OSRM/Яндекс.Карт.", 
             "Первоначальный план предполагал проверку маршрутизации курьеров через запросы к публичным картографическим серверам. В ходе прогона тестов были выявлены сетевые задержки (до 5–8 секунд на тест) и периодические сбои из-за внешних лимитов частоты запросов (HTTP 429). Команда перешла на тестирование алгоритма Дейкстры на локальном графе NetworkX в RAM. Это сократило время прогона тестов с 3 минут до 4 секунд и обеспечило 100% независимость от интернета.")
    p_bullet("• Отклонение 2: Замена изолированного мокирования БД на реальную SQLite in-memory.", 
             "Использование Mock-объектов скрывало ошибки некорректного формирования SQL-запросов и не позволяло тестировать реальные CHECK-ограничения базы данных (например, неотрицательность очереди филиала). Переход на in-memory SQLite позволил выявить 3 критических дефекта схемы на ранней стадии.")

    # 3.3. Оставшиеся тесты и исправления в плане тестирования
    p_h2("3.3. Оставшиеся тесты и исправления в плане тестирования")
    p_body(
        "На момент формирования настоящего отчета все тесты, предусмотренные функциональной спецификацией текущего релиза v1.0, "
        "выполнены на 100%. План тестирования на фазу релиза v1.0 объявляется полностью закрытым."
    )

    p_body(
        "В качестве перспективных задач тестирования для следующего мажорного релиза (Релиз 2.0) сформирован "
        "дополнительный бэклог испытаний:\n"
        "1. Нагрузочное тестирование курьерского шлюза с использованием фреймворка Locust (эмуляция 500+ одновременно активных курьеров, генерирующих GPS-координаты каждые 3 секунды);\n"
        "2. Тестирование WebSocket-соединений для передачи push-уведомлений о переназначении филиалов в режиме реального времени;\n"
        "3. Кроссбраузерное тестирование на мобильных устройствах под управлением iOS Safari и Android Chrome с использованием BrowserStack."
    )

    # =========================================================================
    # 4. ЗОНЫ ТЕСТИРОВАНИЯ
    # =========================================================================
    p_h1("4. Зоны тестирования")

    p_body(
        "Архитектура решения декомпозирована на 6 самостоятельных участков кода — зон тестирования. "
        "Ниже приведены детализированные результаты испытаний по каждой функциональной зоне в формате: "
        "цели, критерии оценивания, фактические результаты и инженерные рекомендации по развитию."
    )

    # Зона 1
    p_h2("4.1. Зона 1: Модуль аутентификации, безопасности и ролевого доступа (Auth & RBAC)")
    p_h3("4.1.1. Цели тестирования")
    p_body("Проверка надежности аутентификации пользователей по стандарту OAuth2 Password Bearer, криптографической стойкости хэширования паролей Bcrypt, валидации JWT-токенов и строгого разграничения прав доступа к эндпоинтам согласно ролям (admin, manager, executive, driver).")
    p_h3("4.1.2. Критерии оценивания")
    p_body("100% блокировка неавторизованных запросов (HTTP 401 Unauthorized); 100% блокировка пользователей при попытке обращения к чужим ролевым функциям (HTTP 403 Forbidden); время генерации и проверки токена < 2 мс; невозможность восстановления открытого пароля из хэша.")
    p_h3("4.1.3. Результаты")
    p_body("Все 25 тестов зоны (test_auth.py) завершены со статусом PASSED. Проверены сценарии регистрации, логина, подделки подписи JWT, истечения времени жизни токена и ролевой изоляции. Уязвимостей не выявлено.")
    p_h3("4.1.4. Рекомендации")
    p_body("В релизе 2.0 рекомендуется реализовать механизм Refresh Tokens с хранением черного списка отозванных токенов в оперативной памяти Redis для мгновенного отзыва сессий скомпрометированных пользователей.")

    # Зона 2
    p_h2("4.2. Зона 2: Модуль агрегации KPI и аналитической статистики заказов (Stats & Analytics)")
    p_h3("4.2.1. Цели тестирования")
    p_body("Верификация векторизованных алгоритмов расчета ключевых показателей эффективности (KPI) на базе Pandas/NumPy: общего числа заказов, среднего времени доставки, процента отмен, нагрузки филиалов и построения таймсерий.")
    p_h3("4.2.2. Критерии оценивания")
    p_body("Математическая точность расчетов (совпадение с эталонными значениями); отсутствие деления на ноль на пустых выборках; непрерывность временных рядов (заполнение пропусков по дням и часам); время вычисления KPI < 100 мс на 10 000 записей.")
    p_h3("4.2.3. Результаты")
    p_body("Все 19 тестов зоны (test_stats.py) пройдены успешно. Алгоритмы устойчивы к краевым условиям (выборка из 0 заказов, 100% отмен, нестандартные диапазоны дат). Время отклика составляет 15–40 мс.")
    p_h3("4.2.4. Рекомендации")
    p_body("Для оптимизации нагрузки при росте исторической базы заказов рекомендуется внедрить агрегационные таблицы в БД (материализованные представления PostgreSQL) с пересчетом раз в сутки.")

    # Зона 3
    p_h2("4.3. Зона 3: Алгоритмический модуль маршрутизации курьеров на графе дорог (Routing & Graphs)")
    p_h3("4.3.1. Цели тестирования")
    p_body("Проверка поиска кратчайшего пути Дейкстры на ориентированном взвешенном графе дорог Нижнего Новгорода, динамического учета коэффициентов пробок и многокритериального выбора филиала с учетом очередей заказов.")
    p_h3("4.3.2. Критерии оценивания")
    p_body("Корректность нахождения оптимального маршрута; перестроение траектории в обход пробок при загрузке новых весов; время расчета пути до всех активных филиалов сети не более 100 мс; сохранение связности графа.")
    p_h3("4.3.3. Результаты")
    p_body("Все 22 теста (test_admin_graph.py, test_recommendations.py) успешно пройдены. При изменении дорожных весов через CSV алгоритм мгновенно адаптирует рекомендации. Среднее время отклика составляет 45–80 мс.")
    p_h3("4.3.4. Рекомендации")
    p_body("Реализовать фоновый демон обновления пробок по расписанию (cron-задача) для автоматической калибровки весов дорог без необходимости ручной загрузки файлов администратором.")

    # Зона 4
    p_h2("4.4. Зона 4: Модуль пространственного машинного обучения K-Means (Demand ML Analysis)")
    p_h3("4.4.1. Цели тестирования")
    p_body("Проверка сходимости алгоритма кластеризации K-Means (Scikit-learn) на геокоординатах доставок, точности определения центроидов спроса, корректности сопоставления с радиусами филиалов и классификации рекомендаций.")
    p_h3("4.4.2. Критерии оценивания")
    p_body("Стабильная сходимость алгоритма за < 30 итераций; генерация рекомендаций трех типов (открыть филиал, закрыть/оптимизировать филиал, зона оптимальна); корректная обработка исключения при N < K; время работы < 300 мс.")
    p_h3("4.4.3. Результаты")
    p_body("Все 10 тестов (test_ml_pipeline.py, test_ml_services.py) пройдены со 100% успехом. Формируемые координаты центроидов точно соответствуют очагам концентрации доставок. Обработка ошибки N < K возвращает HTTP 422.")
    p_h3("4.4.4. Рекомендации")
    p_body("Добавить в модуль алгоритм автоматического определения оптимального числа кластеров (метод «локтя» / Silhouette Score) для подсказки менеджеру наилучшего значения параметра K.")

    # Зона 5
    p_h2("4.5. Зона 5: Конструктор отчетов и потоковый экспорт CSV (Reports & CSV Export)")
    p_h3("4.5.1. Цели тестирования")
    p_body("Верификация CRUD-операций пользовательских и системных шаблонов отчетов, валидации JSON-конфигураций и надежности потоковой генерации файлов CSV с поддержкой кодировки кириллицы.")
    p_h3("4.5.2. Критерии оценивания")
    p_body("Корректное открытие сгенерированного файла CSV в Microsoft Excel и LibreOffice без искажения русских букв (наличие UTF-8 BOM); экранирование кавычек и запятых; отсутствие утечек памяти при потоковой отдаче больших объемов данных.")
    p_h3("4.5.3. Результаты")
    p_body("Все 19 тестов (test_reports.py) завершены успешно. Экспорт CSV протестирован на объемах до 5000 строк. Наличие байтов BOM гарантирует безупречную совместимость с Excel.")
    p_h3("4.5.4. Рекомендации")
    p_body("В релизе 2.0 расширить возможности экспорта поддержкой нативного формата XLSX с форматированием ячеек и генерацией PDF-отчетов с преднастроенными колонтитулами.")

    # Зона 6
    p_h2("4.6. Зона 6: Клиентский интерфейс веб-дашборда и интерактивные карты (Frontend SPA & Maps)")
    p_h3("4.6.1. Цели тестирования")
    p_body("Проверка Single Page Application на Vue 3, реактивного стейт-менеджмента Pinia, ролевой защиты маршрутов Vue Router, интерактивности графиков Chart.js и рендеринга геопространственных карт Leaflet/Folium.")
    p_h3("4.6.2. Критерии оценивания")
    p_body("Отсутствие ошибок JavaScript в консоли браузера; адаптивность интерфейса при разрешениях от 1366×768 до 1920×1080; сокрытие недоступных пунктов навигации в зависимости от роли пользователя.")
    p_h3("4.6.3. Результаты")
    p_body("Клиентское приложение успешно проходит typecheck (TypeScript без ошибок) и production-сборку Vite. Навигационные guard-фильтры надежно блокируют переход водителя в разделы аналитики.")
    p_h3("4.6.4. Рекомендации")
    p_body("Реализовать поддержку темной темы оформления (Dark Mode) и кэширование статических картографических тайлов через Service Workers для ускорения загрузки карт при медленном интернете.")

    # =========================================================================
    # 5. ОТЧЕТ ОБ ОШИБКАХ
    # =========================================================================
    p_h1("5. Отчет об ошибках")

    p_body(
        "Отчет об ошибках систематизирует информацию о дефектах, выявленных и устраненных в ходе испытаний. "
        "Анализ характера и динамики дефектов позволяет объективно оценить текущее качество программного продукта, "
        "стабильность кодовой базы и готовность системы к сдаче заказчику."
    )

    # 5.1. Статус билда
    p_h2("5.1. Статус билда")
    p_body(
        "По результатам финального прогона полного набора автотестов (181 passed) и проведения комплексных приемочных испытаний, "
        "текущему билду системы присвоен статус:"
    )

    p_callout(
        "Идентификатор сборки: Build v1.0.0-rc3 (Release Candidate).\n"
        "Статус стабильности: СТАБИЛЕН (STABLE / PRODUCTION READY).\n"
        "Прохождение тестов: 181 из 181 (100.0% Passed, 0 Failed, 0 Errors).\n"
        "Текущая фаза проекта по MSF: Завершение фазы «Стабилизация решения» (Stabilizing Phase).\n"
        "Решение комиссии по качеству: Программный комплекс готов к релизу и демонстрации Заказчику.",
        title="ИТОГОВЫЙ СТАТУС РЕЛИЗНОГО БИЛДА:",
        bg_color="F0FDF4",
        border_color="16A34A"
    )

    # 5.2. Известные ошибки
    p_h2("5.2. Известные ошибки")
    p_body(
        "В процессе разработки и тестирования было зафиксировано, исследовано и успешно устранено 10 характерных дефектов. "
        "В таблице 5.1 представлен сводный реестр выявленных дефектов с указанием их критичности, симптомов, причин и принятых решений."
    )

    bugs_headers = ["ID", "Модуль", "Критичность", "Симптом дефекта", "Причина ошибки", "Статус и принятое решение"]
    bugs_data = [
        [
            "BUG-01", "StatsService", "Critical",
            "Ошибка деления на ноль (ZeroDivisionError) при расчете среднего времени доставки.",
            "При фильтрации за день отсутствовали доставленные заказы, знаменатель формулы был равен 0.",
            "ЗАКРЫТ. Добавлена проверка len(delivered) > 0 с возвратом 0.0 при пустой выборке."
        ],
        [
            "BUG-02", "ReportService", "Major",
            "Искажение кириллических символов («кракозябры») при открытии CSV в Excel.",
            "Отсутствие метки порядка байтов (BOM) в начале UTF-8 потока приводило к автоопределению CP-1252.",
            "ЗАКРЫТ. В начало потока добавлен заголовок \\xef\\xbb\\xbf (UTF-8 BOM), текст открывается корректно."
        ],
        [
            "BUG-03", "RoutingCore", "Critical",
            "Краш алгоритма Дейкстры (NetworkXNoPath) при запросе маршрута курьером.",
            "Координаты курьера попадали на изолированный тупиковый пешеходный узел графа.",
            "ЗАКРЫТ. Реализован алгоритм предварительного связывания с ближайшей компонентой сильной связности."
        ],
        [
            "BUG-04", "KMeansML", "Major",
            "Аварийное падение Scikit-learn при кластеризации с кодом ошибки ValueError.",
            "Число доступных исторических заказов в выбранном периоде оказалось меньше числа кластеров (N < K).",
            "ЗАКРЫТ. Добавлена валидация входного массива с выбросом доменного исключения и кодом HTTP 422."
        ],
        [
            "BUG-05", "Auth / RBAC", "Major",
            "Курьер мог получить финансовую аналитику сети прямым запросом к API.",
            "В роутере /api/v1/stats/summary отсутствовал защитный фильтр require_roles.",
            "ЗАКРЫТ. Внедрена строгая зависимость require_roles(['admin', 'manager', 'executive']), курьер получает 403."
        ],
        [
            "BUG-06", "StatsTimeseries", "Normal",
            "Разрывы и зависание отрисовки графика динамики заказов на фронтенде.",
            "В дни без заказов Pandas возвращал пропуски дат, нарушая непрерывность таймсерии.",
            "ЗАКРЫТ. Разработан механизм fill_gaps, дополняющий недостающие даты нулевыми значениями."
        ],
        [
            "BUG-07", "DatabaseSeed", "Normal",
            "Ошибка IntegrityError при повторном запуске скрипта наполнения БД seed.py.",
            "Попытка повторной вставки системных шаблонов отчетов с уже существующими именами.",
            "ЗАКРЫТ. Реализован механизм upsert: проверка наличия записей перед созданием (идемпотентность)."
        ],
        [
            "BUG-08", "BranchModel", "Normal",
            "Очередь филиала current_queue принимала отрицательные значения (-3).",
            "Множественные параллельные декременты очереди без валидации нижнего порога.",
            "ЗАКРЫТ. Добавлено CHECK-ограничение ck_branches_current_queue_non_negative на уровне схемы PostgreSQL."
        ],
        [
            "BUG-09", "OrderRepo", "Minor",
            "Расхождения в выборках заказов из-за разных локальных часовых поясов.",
            "Использование naive datetime без информации о временной зоне.",
            "ЗАКРЫТ. Все даты в моделях и API переведены в строгий UTC aware формат с проверкой валидатором."
        ],
        [
            "BUG-10", "FrontendMap", "Trivial",
            "Карта Folium открывалась с масштабом всего мира вместо Нижнего Новгорода.",
            "Некорректный порядок передачи координат (lon, lat вместо lat, lon) в параметрах карты.",
            "ЗАКРЫТ. Координаты нормализованы: центр [56.3269, 44.0075], стартовый зум зафиксирован на 12."
        ]
    ]
    p_table(bugs_headers, bugs_data, [0.7, 1.1, 0.9, 1.7, 1.8, 1.6])

    p_body(
        "На момент стабилизации билда v1.0.0-rc3 все 10 дефектов устранены и верифицированы регрессионными тестами. "
        "Нерешенных критических и блокирующих ошибок (Blokers / Critical) в системе нет."
    )

    # 5.3. Анализ сходимости ошибок
    p_h2("5.3. Анализ сходимости ошибок")
    p_body(
        "В данном подразделе выполняется анализ динамики выявления и устранения дефектов в соответствии с методологией MSF "
        "(Microsoft Solutions Framework). Модель MSF определяет два фундаментальных индикатора завершения стабилизации решения:\n"
        "1. Точка конвергенции ошибок (Bug Convergence);\n"
        "2. Точка достижения нуля ошибок (Zero-Bug Bounce)."
    )

    p_body(
        "В точке конвергенции (Bug Convergence) становится заметен устойчивый прогресс в стабилизации системы: "
        "скорость устранения ошибок разработчиками начинает стабильно превосходить скорость их обнаружения тестировщиками. "
        "График количества активных дефектов проходит свой максимум и переходит в фазу неуклонного спада."
    )

    p_image("docs/diagrams/bug_convergence.png", "Рис. 5.1. Точка конвергенции ошибок (Bug Convergence). Источник: Модель процессов MSF", width_in=4.2)

    p_body(
        "Поскольку количество найденных, но не устраненных ошибок может кратковременно колебаться из-за углубленного тестирования "
        "новых модулей, конвергенция рассматривается как тенденция. Для команды «NewFastPizza» точка конвергенции была официально "
        "зафиксирована 15 декабря 2026 года (конец 6-й недели), когда количество закрываемых багов устойчиво превысило приток новых."
    )

    p_body(
        "Точка достижения нуля (Zero-Bug Bounce) — это момент времени, когда впервые в ходе проекта абсолютно все выявленные "
        "ошибки оказываются устраненными (активный бэклог дефектов равен 0). Вслед за этой точкой пики новых ошибок становятся "
        "незначительными, вплоть до полного угасания перед выпуском релизного кандидата."
    )

    p_image("docs/diagrams/zero_bug_bounce.png", "Рис. 5.2. Точка достижения нуля (Zero-Bug Bounce). Источник: Модель процессов MSF", width_in=4.6)

    p_body(
        "В проекте «NewFastPizza» точка достижения нуля (Zero-Bug Bounce) была достигнута 22 декабря 2026 года "
        "(в начале 8-й недели разработки). Последующие контрольные прогоны выявили лишь одну косметическую ошибку центрирования карты "
        "(BUG-10), которая была оперативно устранена, что обеспечило идеальное затухание графика дефектов к моменту финальной сдачи."
    )

    p_body("В таблице 5.2 приведена понедельная статистика выявления и устранения ошибок проекта:")

    trends_headers = ["Неделя проекта", "Период", "Выявлено новых", "Устранено", "Активных ошибок", "Событие / Веха MSF"]
    trends_data = [
        ["Неделя 1–2", "01.11–14.11.2026", "0", "0", "0", "Проектирование и ТЗ. Код не пишется."],
        ["Неделя 3", "15.11–21.11.2026", "2", "1", "1", "Старт кодирования ORM. Обнаружен BUG-08."],
        ["Неделя 4", "22.11–28.11.2026", "3", "2", "2", "Тестирование графов и KPI. BUG-01, BUG-03."],
        ["Неделя 5", "29.11–05.12.2026", "3", "2", "3", "Интеграция K-Means и CSV. Пик дефектов."],
        ["Неделя 6", "06.12–12.12.2026", "1", "3", "1", "ТОЧКА КОНВЕРГЕНЦИИ (15.12.2026)."],
        ["Неделя 7", "13.12–19.12.2026", "1", "1", "1", "Стабилизация ролей RBAC. Устранен BUG-05."],
        ["Неделя 8", "20.12–26.12.2026", "0", "1", "0", "ТОЧКА ДОСТИЖЕНИЯ НУЛЯ (Zero-Bug Bounce)."]
    ]
    p_table(trends_headers, trends_data, [1.1, 1.5, 1.1, 1.0, 1.1, 2.0])

    # =========================================================================
    # 6. ОБНОВЛЕННЫЙ ПЛАН ТЕСТИРОВАНИЯ
    # =========================================================================
    p_h1("6. Обновленный план тестирования")

    p_body(
        "Раздел обобщает изменения и дополнения, внесенные в план тестирования по итогам фаз разработки и стабилизации. "
        "В связи с успешным достижением критериев одобрения решения и полным закрытием бэклога дефектов, акцент тестирования "
        "смещается с поиска программных ошибок на регламентные эксплуатационные процедуры."
    )

    p_h2("Ключевые обновления плана тестирования:")
    p_bullet("1. Включение постоянного Smoke-тестирования в регламент развертывания:", 
             "Автоматизированный скрипт scripts/smoke-test.sh утвержден как обязательный этап при каждом развертывании или обновлении на production VPS сервере. Скрипт опрашивает системный health-эндпоинт и подтверждает готовность базы данных и дорожного графа перед переключением пользовательского трафика.")
    p_bullet("2. Фиксация набора регрессионных тестов (181 тест):", 
             "Все разработанные модульные и интеграционные тесты зафиксированы в качестве постоянного регрессионного шлюза (Regression Test Suite). Любые последующие доработки системы должны сопровождаться 100% прохождением данного набора.")
    p_bullet("3. Утверждение методики приемо-сдаточных испытаний (UAT):", 
             "Сформирован демонстрационный сценарий для сдачи проекта Заказчику (преподавателю), включающий: вход под разными ролями пользователей, построение графиков KPI, фильтрацию и экспорт отчета в CSV, генерацию тепловой карты спроса K-Means и расчет маршрута курьера в реальном времени.")
    p_bullet("4. План нагрузочного тестирования на релиз 2.0:", 
             "Запланировано проведение нагрузочных испытаний курьерского шлюза с симуляцией 500+ активных водителей и исследование деградации времени отклика алгоритма Дейкстры при экстремальных объемах дорожной сети.")

    # Save final document
    out_dir = "docs"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "5. Тестирование и отчеты об ошибках.docx")
    doc.save(out_path)
    print(f"Document 5 successfully created and saved to: {out_path}")
    return out_path

if __name__ == '__main__':
    create_document()


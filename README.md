import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_element(name):
    return OxmlElement(name)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

doc = docx.Document()

# Установка полей страницы (2 см со всех сторон)
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.79)
    section.bottom_margin = Inches(0.79)
    section.left_margin = Inches(0.79)
    section.right_margin = Inches(0.79)

title = doc.add_paragraph()
run = title.add_run("Дорожная карта проекта")
run.bold = True
run.font.name = 'Arial'
run.font.size = Pt(14)
title.alignment = WD_ALIGN_PARAGRAPH.LEFT

# Таблица: 11 строк, 5 столбцов
table = doc.add_table(rows=11, cols=5)
table.style = 'Table Grid'

headers = ["№", "Комплекс мер", "Место проведения", "Срок проведения", "Примечания"]
col_widths = [Inches(0.5), Inches(2.5), Inches(1.5), Inches(1.2), Inches(1.3)]

# Заполнение заголовка
hdr_cells = table.rows[0].cells
for i, text in enumerate(headers):
    hdr_cells[i].text = text
    set_cell_margins(hdr_cells[i])
    for p in hdr_cells[i].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.name = 'Arial'
            r.font.size = Pt(11)

# Данные для таблицы
data = [
    # Строка 1: Этап 1 (Объединенная)
    {"is_stage": True, "text": "Этап 1"},
    {"is_stage": False, "vals": ["1", "Анализ существующих аналогов вычислительных систем и ПО", "Дом / Цифровая среда", "Сентябрь", "Индивидуальная исследовательская деятельность"]},
    {"is_stage": False, "vals": ["2", "Изучение содержания: разбор алгоритма сортировочной станции Э. Дейкстры", "Школьная медиатека / Дом", "Сентябрь", "Консультация с научным руководителем"]},
    {"is_stage": False, "vals": ["3", "Выбор темы: «Разработка интерпретатора математических выражений на С++»", "Школа", "Сентябрь", "Утверждено руководителем"]},
    {"is_stage": False, "vals": ["4", "Разработка базового ядра программы (лексический и синтаксический анализ)", "Дом", "Сентябрь", "Учитель (контроль)"]},
    # Этап 2
    {"is_stage": True, "text": "Этап 2"},
    {"is_stage": False, "vals": ["5", "Создание программного продукта (реализация движка вычислений и GUI-оболочки)", "Гугл документы...", "Ноябрь", "Паспорт проекта, исходный код"]},
    {"is_stage": False, "vals": ["6", "Участие в конкурсе / школьной научно-практической конференции", "Школа", "Сентябрь - октябрь", "Предварительная защита"]},
    {"is_stage": False, "vals": ["7", "Апробирование калькулятора и методических заданий в восьмой параллели", "Школа (уроки информатики)", "Октябрь - май", "Проверка работы ПО на уроках"]},
    # Этап 3
    {"is_stage": True, "text": "Этап 3"},
    {"is_stage": False, "vals": ["8", "Анализ результатов апробирования, сбор отзывов от учащихся 8 классов", "Школа", "Май", "Статистика ошибок"]},
    {"is_stage": False, "vals": ["9", "Оценка эффективности проекта и финальная защита", "Школа", "Май", "Выставление итоговой оценки"]}
]

row_idx = 1
for row_data in data:
    row_cells = table.rows[row_idx].cells
    if row_data.get("is_stage"):
        # Объединяем ячейки для этапа
        merged = row_cells[0].merge(row_cells[4])
        merged.text = row_data["text"]
        set_cell_margins(merged)
        # Делаем серый фон для строки этапа
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="F2F2F2"/>'
        merged._tc.get_or_add_tcPr().append(parse_xml(shading_xml))
        for p in merged.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.name = 'Arial'
                r.font.size = Pt(11)
    else:
        vals = row_data["vals"]
        for col_idx, text in enumerate(vals):
            row_cells[col_idx].text = text
            set_cell_margins(row_cells[col_idx])
            for p in row_cells[col_idx].paragraphs:
                if col_idx == 0 or col_idx == 3:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = 'Arial'
                    r.font.size = Pt(10)
    row_idx += 1

# Принудительная установка ширины колонок
for row in table.rows:
    for idx, width in enumerate(col_widths):
        row.cells[idx].width = width

import os
os.makedirs("generated", exist_ok=True)
doc.save("generated/Project_Roadmap.docx")

import pymupdf
import re
from copy import copy
from io import BytesIO

from openpyxl import load_workbook


def pdf_spec_to_row_list(pdf_path):
    """Извлекает из pdf-файла спецификации список строк спецификации."""
    with pymupdf.open(pdf_path) as doc:
        return parse_spec(doc)


def parse_spec(doc):
    spec = []
    prev_spec_line = None
    first_table = True
    for page in doc:
        tabs = page.find_tables()  # locate and extract any tables on page
        if not tabs:
            continue
        for tab in tabs:
            in_spec = False  # не дошёл до спецификации
            lines = tab.extract()
            for line in lines:
                #print(line)
                if not in_spec:
                    if "Примечание" in line or "Примечания" in line:  # шапка таблицы спецификации
                        spec_line_cols_count = len(line)  # количество столбцов в pdf-таблице спецификации
                        pattern = detect_pattern(line)  # шаблон таблицы спецификации
                        spec_col_count = len(pattern)
                        if spec_col_count != 9:
                            raise ValueError(f"В спецификации должно быть 9 столбцов, а не {spec_col_count}")
                        if first_table:  # только для первой шапки таблицы
                            spec.append(gost_spec_title())  # добавляем в спецификацию шапку по ГОСТ 21.110-2013
                            pass
                        in_spec = True # внутри спецификации
                        first_table = False
                        continue
                if in_spec:
                    # if None in line[first_index: last_index+1]: break
                    if len(line) != spec_line_cols_count or None in [line[i] for i in pattern]: continue  # игнорируем строки, набор столбцов которых не соответствует ранее зафиксированному набору для мпецификации
                    spec_line = [line[i] for i in pattern]
                    if all(cell == '' for cell in spec_line): continue  # Все столбцы содержат ''
                    normalize_row(spec_line, prev_spec_line)
                    if spec_line == ['1', '2', '3', '4', '5', '6', '7', '8', '9']: continue  # ['1', '2', '3', '4', '5', '6', '7', '8', '9'] игнорируем
                    spec.append(spec_line)
                    prev_spec_line = spec_line
                    # spec.append(normalize_row(line[first_index: last_index+1]))
    if spec == []:
        raise ValueError("В файле спецификация не найдена")
    return spec


def detect_pattern(line):
    """Определяет шаблон таблицы спецификации по шапке спецификации (в которой все поля не пустые).
    input: line - список столбцов pdf-таблицы спецификации
    output: pattern - список индексов столбцов, содерщащих данные спецификации.
    """
    pattern = []
    for i, col in enumerate(line):
        if col == '' or col is None:
            continue
        pattern.append(i)
    return pattern
    # return [i for i, col in enumerate(line) if col not in ('', None)] # быстрее на 10–30%


def normalize_row(spec_line, prev_spec_line):
    """Очищает строку: заменяет удаляет лишние пробелы."""
    for col in spec_line:
        if not isinstance(col, str):
            continue
        # Замена символов перевода строки на пробел
        # text = text.replace('\n', ' ')
        # Удаление множественных пробелов (один или более пробелов заменяем на один)
        col = re.sub(r' +', ' ', col)
        # Удаление пробелов в начале и конце (опционально)
        col = col.strip()
        # if col == '': col = "—"  # замена '' на длинное тире (для нейросетей)
    # ставим пробел при ГОСТ12345-67
    spec_line[2] = re.sub(r'ГОСТ(\d)', r'ГОСТ \1', spec_line[2])
    # Приводим единицу измерения к стандартному обозначению
    spec_line[5] = normalize_unit(spec_line[5])
    # Меняем точку на запятую в Кол-во и Ед. масса
    spec_line[6].replace('.', ',')
    spec_line[7].replace('.', ',')
    if spec_line[1].lower() == "то же":
         spec_line[1] = prev_spec_line[1]



def gost_spec_title():
     """Заголовок таблицы спецификации по ГОСТ 21.110-2013"""
     spec_title = ["Поз.",
                   "Наименование и техническая характеристика",
                   "Тип, марка, обозначение документа, опросного листа",
                   "Код продукции",
                   "Поставщик",
                   "Ед. измерения",
                   "Количество",
                   "Масса 1 ед., кг",
                   "Примечание"]
     return spec_title


def row_list_to_xlsx_bytes(spec, pdf_stem, template_path):
    """Формирует xlsx-файл спецификации на базе шаблона и возвращает его в виде байтов."""
    if not template_path.exists():
        raise FileNotFoundError(f"Шаблон спецификации не найден: {template_path}")

    wb = load_workbook(template_path)
    ws = wb.active
    if ws is None:
        raise ValueError("Не удалось получить активный лист из шаблона")

    # Записываем имя PDF в ячейку A1
    ws["A1"] = pdf_stem
    
    # 1. Сохраняем референсные стили из строки 3 (A3:I3)
    # Именно здесь в шаблоне настроены перенос текста, выравнивание и границы
    ref_styles = {}
    for col_idx in range(1, 10): # Столбцы A-I (индексы 1-9)
        ref_cell = ws.cell(row=3, column=col_idx)
        ref_styles[col_idx] = {
            'font': copy(ref_cell.font),
            'border': copy(ref_cell.border),
            'fill': copy(ref_cell.fill),
            'alignment': copy(ref_cell.alignment),
            'number_format': ref_cell.number_format
        }

    # 2. Пропускаем первую строку spec (так как это заголовок) 
    # и начинаем вывод данных с 3-й строки Excel
    data_rows = spec[1:] if len(spec) > 1 else []
    
    # 3. Заполняем данные и применяем сохраненные стили
    for row_offset, row_values in enumerate(data_rows):
        row_idx = 3 + row_offset
        for col_idx, value in enumerate(row_values, start=1):
            if col_idx > 9:
                break
                
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            
            # Применяем скопированные стили из эталонной строки
            styles = ref_styles.get(col_idx)
            if styles:
                cell.font = styles['font']
                cell.border = styles['border']
                cell.fill = styles['fill']
                cell.alignment = styles['alignment']
                cell.number_format = styles['number_format']

    # 4. Жестко фиксируем автофильтр на строке заголовков (строка 2)
    # Указываем диапазон до последней заполненной строки, чтобы фильтр не "уехал"
    last_row = 2 + len(data_rows)
    if last_row > 2:
        ws.auto_filter.ref = f"A2:I{last_row}"
    else:
        ws.auto_filter.ref = "A2:I2"

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def row_list_to_md(lines):
    """Преобразует список строк в Markdown-таблицу."""
    md = ""
    for line in lines:
        md += "| " + " | ".join(line) + " |\n"
    return md

import re

def normalize_unit(unit: str) -> str:
    """
    Приводит единицу измерения к стандартному обозначению.
    Все варианты (с точками, пробелами, разным регистром) нормализуются
    путём удаления разделителей и приведения к нижнему регистру.
    """
    # Приводим к нижнему регистру и убираем лишние пробелы
    unit = unit.lower().strip()
    # Заменяем Unicode-индексы на обычные цифры
    unit = unit.replace('²', '2').replace('³', '3')
    # Удаляем все точки и пробелы (оставляем только буквы и цифры)
    cleaned = re.sub(r'[.\s]+', '', unit)
    
    # Словарь: нормализованный ключ -> стандартное обозначение
    MAPPING = {
        # Метры
        'м': 'м',
        'метр': 'м',
        'метры': 'м',
        'm': 'м',
        'metr': 'м',
        # Квадратные метры
        'м2': 'м2',
        'м^2': 'м2',
        'квм': 'м2',          # кв.м, кв м, кв.м. – все сводятся к "квм"
        'квадратныйметр': 'м2',
        'квадратныеметры': 'м2',
        'квадратныхметров': 'м2',
        'sqm': 'м2',
        # Кубические метры
        'м3': 'м3',
        'м^3': 'м3',
        'кубм': 'м3',
        'кубическийметр': 'м3',
        'кубическиеметры': 'м3',
        'кубическихметров': 'м3',
        'cum': 'м³',  #почему-то выдаёт ошибку
        # Литры
        'л': 'л',
        'литр': 'л',
        'литры': 'л',
        'литров': 'л',
        'l': 'л',
        # Килограммы
        'кг': 'кг',
        'килограмм': 'кг',
        'kg': 'кг',
        # Тонны
        'т': 'т',
        'тн': 'т',
        'тонны': 'т',
        'тонна': 'т',
        't': 'т',
        # Штуки
        'шт': 'шт.',
        'штука': 'шт.',
        'штук': 'шт.',
        # Упаковки
        'уп': 'уп.',
        'упаковка': 'уп.',
        'упаковки': 'уп.',
        'упаковок': 'уп.',
        # Вёдра
        'вед': 'вед.',
        'ведро': 'вед.',
        'ведра': 'вед.',
        'ведер': 'вед.',
        'вёдер': 'вед.',
        # Баллоны
        'бал': 'бал.',
        'баллон': 'бал.',
        'баллоны': 'бал.',
        'баллонов': 'бал.',
        # Погонные метры
        'погм': 'пог. м',
        'пм': 'пог. м',
        'погонныйметр': 'пог. м',
        'погонныеметры': 'пог. м',
        'погонныхметров': 'пог. м',
    }
    
    # Ищем в словаре по очищенному ключу
    if cleaned in MAPPING:
        return MAPPING[cleaned]
    
    # Если не найдено – возвращаем исходную строку (или можно вернуть cleaned)
    return unit
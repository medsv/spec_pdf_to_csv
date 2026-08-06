import pymupdf
import re


def spec_pdf_to_row_list(pdf_path):
    """Извлекает из pdf-файла спецификации список строк спецификации."""
    with pymupdf.open(pdf_path) as doc:
        return parse_spec(doc)


def parse_spec(doc):
    spec = []
    first_table = True
    for page in doc:
        tabs = page.find_tables()  # locate and extract any tables on page
        if not tabs:
            continue
        for tab in tabs:
            in_spec = False  # не дошёл до спецификации
            lines = tab.extract()
            for line in lines:
                # print(line)
                if not in_spec:
                    if "Примечание" in line or "Примечания" in line:  # шапка таблицы спецификации
                        spec_line_cols_count = len(line)  # количество столбцов в pdf-таблице спецификации
                        pattern = detect_pattern(line)  # шаблон таблицы спецификации
                        spec_col_count = len(pattern)
                        if spec_col_count != 9:
                            raise ValueError(f"В спецификации должно быть 9 столбцов, а не {spec_col_count}")
                        if first_table:  # только для первой шапки таблицы
                            spec.append(gost_spec_title())  # добавляем в спецификацию шапку по ГОСТ 21.110-2013
                        in_spec = True
                        first_table = False
                        continue
                if in_spec:
                    # if None in line[first_index: last_index+1]: break
                    if len(line) != spec_line_cols_count or None in [line[i] for i in pattern]: continue  # игнорируем строки, набор столбцов которых не соответствует ранее зафиксированному набору для мпецификации
                    spec_line = [clean_text(line[i]) for i in pattern]
                    if spec_line == ['1', '2', '3', '4', '5', '6', '7', '8', '9']: continue  # ['1', '2', '3', '4', '5', '6', '7', '8', '9'] игнорируем
                    spec.append(spec_line)
                    # spec.append(clean_text(line[first_index: last_index+1]))
    if spec == []:
        raise ValueError("В файле спецификация не найдена")
    return spec


def detect_pattern(line):
    """Определяет шаблон таблицы спецификации.
    input: line - список столбцов pdf-таблицы спецификации
    output: pattern - список индексов столбцов, содерщащих данные спецификации.
    """
    pattern = []
    for i, col in enumerate(line):
        if col == '' or col is None:
            continue
        pattern.append(i)
    return pattern
    # return [i for i, col in enumerate(line) if col not in ('', None)] # быстрее на 10–30 %


def clean_text(text):
    """Очищает строку: заменяет \\n на пробел, удаляет лишние пробелы."""
    if not isinstance(text, str):
        return text
    # Замена символов перевода строки на пробел
    # text = text.replace('\n', ' ')
    # Удаление множественных пробелов (один или более пробелов заменяем на один)
    text = re.sub(r' +', ' ', text)
    # Удаление пробелов в начале и конце (опционально)
    text = text.strip()
    return text


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

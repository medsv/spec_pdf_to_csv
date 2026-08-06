import streamlit as st
import os
import tempfile
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# Импортируем вашу библиотеку
from libs.spec_pdf_to_csv import spec_pdf_to_row_list

# Настройки страницы
st.set_page_config(
    page_title="Конвертер спецификаций PDF в XLSX",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Конвертер спецификаций рабочей документации (PDF -> EXLS)")
st.markdown("Загрузите PDF-файл спецификации для её перевода в формат EXLS.")

# Путь к шаблону спецификации
TEMPLATE_PATH = Path("templates") / "Шаблон_спецификации_РД.xlsx"


def spec_to_xlsx(spec, pdf_file_name):
    """Формирует xlsx-файл спецификации на базе шаблона.
    spec: список строк спецификации (список списков).
    pdf_file_name: имя исходного pdf-файла (строка), используется для имени результата.
    """
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Шаблон спецификации не найден: {TEMPLATE_PATH}")

    output_path = Path(pdf_file_name).with_suffix(".xlsx")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    if ws is None:
        raise ValueError("Не удалось получить активный лист из шаблона")

    ws["A1"] = Path(pdf_file_name).stem
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, row_values in enumerate(spec, start=3):
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

    wb.save(output_path)
    return output_path


# 1. Загрузка файла
uploaded_file = st.file_uploader(
    "Выберите PDF файл или мышкой перетащите его сюда", type=["pdf"]
)

if uploaded_file is not None:
    st.info(f"Загружен файл: **{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} КБ)")

    tmp_pdf_path = None
    tmp_xlsx_path = None

    # Создаем временный PDF-файл
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(uploaded_file.getvalue())
        tmp_pdf_path = tmp_pdf.name

    # 2. Кнопка запуска обработки
    if st.button("🚀 Извлечь данные и сформировать EXLS", type="primary"):
        with st.spinner("Идёт анализ PDF ... Пожалуйста, подождите."):
            try:
                spec = spec_pdf_to_row_list(tmp_pdf_path)
                output_path = spec_to_xlsx(spec, uploaded_file.name)
                tmp_xlsx_path = output_path
                st.success("✅ Файл успешно обработан!")

                # 3. Кнопка скачивания
                with open(output_path, "rb") as f:
                    xlsx_bytes = f.read()

                st.download_button(
                    label="📥 Скачать Excel файл",
                    data=xlsx_bytes,
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

            except Exception as e:
                st.error(f"❌ Произошла ошибка при обработке файла: {e}")
                st.exception(e)

            finally:
                # 4. Очистка временных файлов
                for path in (tmp_pdf_path, tmp_xlsx_path):
                    if path:
                        try:
                            os.remove(path)
                        except OSError:
                            pass

else:
    st.warning("Пожалуйста, загрузите PDF файл для начала работы.")

st.markdown(
    """
    <hr>
    <p style="text-align: left; color: gray;">
    <small>
    2026, С.В. Медведев, engpython@yandex.ru
    </small>
    </p>
    """,
    unsafe_allow_html=True,
)

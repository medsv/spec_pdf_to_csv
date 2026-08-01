import streamlit as st
import os
import tempfile

# Импортируем вашу библиотеку
from libs.spec_pdf_to_csv import spec_pdf_to_csv

# Настройки страницы
st.set_page_config(
    page_title="Конвертер спецификаций PDF в CSV",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Конвертер спецификаций рабочей документации (PDF -> CSV)")
st.markdown("Загрузите PDF-файл спецификации для её перевода в формат CSV.")

# 1. Загрузка файла
uploaded_file = st.file_uploader("Выберите PDF файл или мышкой перетащите его сюда", type=["pdf"])

if uploaded_file is not None:
    st.info(f"Загружен файл: **{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} КБ)")
    
    # Создаем временный PDF-файл
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(uploaded_file.getvalue())
        tmp_pdf_path = tmp_pdf.name
        
    # Путь к CSV формируется автоматически: то же имя, расширение .csv
    tmp_csv_path = os.path.splitext(tmp_pdf_path)[0] + ".csv"

    # 2. Кнопка запуска обработки
    if st.button("🚀 Извлечь данные и сформировать CSV", type="primary"):
        
        with st.spinner("Идет анализ PDF с помощью PyMuPDF... Пожалуйста, подождите."):
            try:
                # ==========================================
                # ВЫЗОВ ВАШЕЙ БИБЛИОТЕКИ
                # ==========================================
                # Ваша библиотека сохраняет CSV рядом с PDF 
                # (с тем же именем, но расширением .csv),
                # поэтому достаточно передать путь к PDF.
                spec_pdf_to_csv(tmp_pdf_path)
                # ==========================================

                st.success("✅ Файл успешно обработан!")

                # 3. Кнопка скачивания
                with open(tmp_csv_path, "rb") as file:
                    csv_data = file.read()
                    
                base_name = os.path.splitext(uploaded_file.name)[0]
                download_name = f"{base_name}.csv"
                
                st.download_button(
                    label="📥 Скачать CSV файл",
                    data=csv_data,
                    file_name=download_name,
                    mime="text/csv",
                    type="primary"
                )



            except Exception as e:
                st.error(f"❌ Произошла ошибка при обработке файла: {e}")
                st.exception(e)
                
            finally:
                # 4. Очистка временных файлов
                for path in (tmp_pdf_path, tmp_csv_path):
                    if os.path.exists(path):
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
        unsafe_allow_html=True
    )

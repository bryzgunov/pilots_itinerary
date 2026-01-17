import streamlit as st
import os
import tempfile
from your_script import process_two_pdfs
from datetime import datetime

# Настройки страницы
st.set_page_config(
    page_title="Обработчик двух PDF файлов",
    page_icon="📑",
    layout="centered"
)

# Стили
st.markdown("""
<style>
    .stButton > button {
        background-color: #4CAF50;
        color: white;
        font-size: 18px;
        height: 50px;
        width: 100%;
        border-radius: 10px;
        border: none;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
    .success-msg {
        background-color: #d4edda;
        color: #155724;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #c3e6cb;
        margin: 10px 0;
    }
    .error-msg {
        background-color: #f8d7da;
        color: #721c24;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #f5c6cb;
        margin: 10px 0;
    }
    .info-box {
        background-color: #e7f3fe;
        color: #0c5460;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #b8daff;
        margin: 10px 0;
    }
    .file-info {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #dee2e6;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.title("📑 Обработчик двух PDF файлов")
st.markdown("Загрузите два PDF файла: один с данными Takeoff, другой с основным маршрутом")

# Информация о формате
st.markdown("""
<div class="info-box">
<h4>📌 Требования к файлам:</h4>
<ul>
<li><b>Файл 1:</b> PDF с данными Takeoff (должен содержать "Takeoff" в начале)</li>
<li><b>Файл 2:</b> PDF с основным маршрутом (должен содержать таблицу маршрута)</li>
<li>Система автоматически определит, какой файл содержит Takeoff</li>
<li>На выходе — один Excel файл с двумя листами</li>
</ul>
</div>
""", unsafe_allow_html=True)

# Загрузка двух файлов
st.subheader("📤 Загрузка файлов")

col1, col2 = st.columns(2)

with col1:
    uploaded_file1 = st.file_uploader(
        "Первый PDF файл",
        type=['pdf'],
        help="PDF файл (либо Takeoff, либо основной маршрут)"
    )

with col2:
    uploaded_file2 = st.file_uploader(
        "Второй PDF файл", 
        type=['pdf'],
        help="PDF файл (второй из пары)"
    )

# Отображение информации о загруженных файлах
if uploaded_file1 and uploaded_file2:
    st.markdown("---")
    st.subheader("📋 Загруженные файлы")
    
    # Информация о файлах
    file_info_col1, file_info_col2 = st.columns(2)
    
    with file_info_col1:
        st.markdown(f"""
        <div class="file-info">
        <h4>Файл 1:</h4>
        <p><b>Имя:</b> {uploaded_file1.name}</p>
        <p><b>Размер:</b> {uploaded_file1.size / 1024:.1f} KB</p>
        </div>
        """, unsafe_allow_html=True)
    
    with file_info_col2:
        st.markdown(f"""
        <div class="file-info">
        <h4>Файл 2:</h4>
        <p><b>Имя:</b> {uploaded_file2.name}</p>
        <p><b>Размер:</b> {uploaded_file2.size / 1024:.1f} KB</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Проверка на одинаковые имена
    if uploaded_file1.name == uploaded_file2.name:
        st.error("❌ Ошибка: Файлы имеют одинаковые имена. Пожалуйста, загрузите разные файлы.")
    else:
        # Кнопка обработки
        st.markdown("---")
        if st.button("🚀 Начать обработку файлов", type="primary", use_container_width=True):
            try:
                with st.spinner("⏳ Обработка файлов... Это может занять несколько секунд"):
                    # Получаем байты файлов
                    file1_bytes = uploaded_file1.getvalue()
                    file2_bytes = uploaded_file2.getvalue()
                    
                    # Обрабатываем файлы
                    excel_bytes = process_two_pdfs(
                        file1_bytes, 
                        file2_bytes,
                        uploaded_file1.name,
                        uploaded_file2.name
                    )
                
                # Успешное завершение
                st.markdown('<div class="success-msg">✅ Файлы успешно обработаны!</div>', unsafe_allow_html=True)
                
                # Генерируем имя выходного файла
                output_filename = f"Flight_Log_Extracted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                # Кнопка скачивания
                st.download_button(
                    label=f"⬇️ Скачать Excel файл: {output_filename}",
                    data=excel_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
                # Информация о содержимом
                st.info("""
                **Содержимое Excel файла:**
                - **Лист 'Основное':** Ключевые данные из начала документа
                - **Лист 'Main_Route_Grid':** Распарсенная таблица маршрута
                """)
                
                # Анимация успеха
                st.balloons()
                
            except Exception as e:
                st.markdown(f'<div class="error-msg">❌ Ошибка обработки: {str(e)}</div>', unsafe_allow_html=True)
                
                # Кнопка для повторной попытки
                if st.button("🔄 Попробовать снова", type="secondary"):
                    st.rerun()

# Боковая панель с информацией
with st.sidebar:
    st.header("ℹ️ О сервисе")
    
    st.markdown("""
    ### Как это работает:
    1. **Загрузите два PDF файла:**
       - Один с данными Takeoff
       - Другой с основным маршрутом
    2. **Система определит** автоматически, какой файл содержит Takeoff
    3. **Обработает основной файл** и извлечет данные
    4. **Создаст Excel файл** с результатами
    
    ### Технологии:
    - **PyMuPDF** для чтения PDF
    - **Pandas** для обработки данных
    - **OpenPyXL** для создания Excel
    
    ### Поддержка:
    - Только PDF файлы
    - До 50MB на файл
    - Автоопределение Takeoff
    """)
    
    # Проверка наличия скрипта
    st.markdown("---")
    if os.path.exists("your_script.py"):
        st.success("✅ Скрипт обработки найден")
    else:
        st.error("❌ Скрипт your_script.py не найден")
    
    # Информация о версии
    st.markdown("---")
    st.caption(f"Время: {datetime.now().strftime('%H:%M:%S')}")
    st.caption("v2.0 | Обработка двух файлов")

# Футер
st.markdown("---")
st.caption("Обработчик PDF файлов маршрутных листов | Создано с Streamlit")

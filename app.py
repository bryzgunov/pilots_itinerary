import streamlit as st
import os
import tempfile
import sys
import importlib.util
import shutil
import time
from datetime import datetime

# Настройки страницы
st.set_page_config(
    page_title="PDF Парсер маршрутных листов",
    page_icon="📋",
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
</style>
""", unsafe_allow_html=True)

# Заголовок
st.title("📋 Парсер PDF маршрутных листов")
st.markdown("Загрузите PDF файл маршрутного листа для автоматического парсинга в Excel")

# Информация о поддерживаемом формате
st.markdown("""
<div class="info-box">
<h4>📌 Поддерживаемый формат PDF:</h4>
<ul>
<li>Маршрутные листы авиакомпаний</li>
<li>Должна присутствовать строка заголовка: "WAYPOINT AIRWAY HDG CRS ALT CMP DIR/SPD..."</li>
<li>Только первая страница документа</li>
</ul>
</div>
""", unsafe_allow_html=True)

# Функция для импорта вашего скрипта
def import_my_script():
    """Пытается импортировать ваш скрипт обработки"""
    script_name = "your_script.py"
    
    if not os.path.exists(script_name):
        st.warning(f"⚠️ Файл {script_name} не найден")
        return None
    
    try:
        spec = importlib.util.spec_from_file_location("my_script", script_name)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        st.success(f"✅ Скрипт {script_name} успешно загружен")
        return module
    except Exception as e:
        st.error(f"❌ Ошибка загрузки скрипта: {str(e)}")
        return None

# Функция обработки PDF файла
def process_pdf_file(uploaded_file):
    """Обрабатывает загруженный PDF файл"""
    
    # Создаем прогресс-бар
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Шаг 1: Сохранение PDF файла
    status_text.text("🔄 Сохраняю PDF файл...")
    progress_bar.progress(10)
    
    # Создаем временный файл для PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
        tmp_pdf.write(uploaded_file.getvalue())
        input_pdf_path = tmp_pdf.name
    
    # Шаг 2: Подготовка выходного Excel файла
    status_text.text("🔄 Подготавливаю Excel файл...")
    progress_bar.progress(20)
    
    # Генерируем имя выходного файла
    original_name = os.path.splitext(uploaded_file.name)[0]
    output_excel_name = f"{original_name}_parsed.xlsx"
    output_excel_path = os.path.join(tempfile.gettempdir(), output_excel_name)
    
    # Шаг 3: Импорт и выполнение вашего скрипта
    status_text.text("🔄 Загружаю парсер PDF...")
    progress_bar.progress(30)
    
    try:
        # Импортируем ваш скрипт
        my_script = import_my_script()
        
        if my_script is None:
            raise Exception("Скрипт обработки не найден или не загружен")
        
        # Проверяем наличие функции process
        if not hasattr(my_script, 'process'):
            raise Exception("В скрипте не найдена функция 'process(input_path, output_path)'")
        
        # Шаг 4: Выполнение обработки
        status_text.text("🔄 Парсю PDF файл...")
        progress_bar.progress(50)
        
        # Вызываем вашу функцию обработки
        success = my_script.process(input_pdf_path, output_excel_path)
        
        if not success:
            raise Exception("Обработка завершилась неудачно")
        
        status_text.text("🔄 Форматирую результат...")
        progress_bar.progress(80)
        
        # Проверяем, создан ли выходной файл
        if not os.path.exists(output_excel_path):
            raise Exception("Выходной Excel файл не создан")
        
        # Читаем результат
        with open(output_excel_path, 'rb') as f:
            excel_data = f.read()
        
        progress_bar.progress(100)
        status_text.text("✅ Обработка завершена!")
        time.sleep(0.5)
        
        # Очистка
        progress_bar.empty()
        status_text.empty()
        
        # Удаляем временные файлы
        try:
            os.unlink(input_pdf_path)
            os.unlink(output_excel_path)
        except:
            pass
        
        return excel_data, output_excel_name
    
    except Exception as e:
        # Очистка при ошибке
        progress_bar.empty()
        status_text.empty()
        
        # Удаляем временные файлы
        try:
            if os.path.exists(input_pdf_path):
                os.unlink(input_pdf_path)
            if os.path.exists(output_excel_path):
                os.unlink(output_excel_path)
        except:
            pass
        
        raise e

# Основной интерфейс
st.markdown("---")

# Загрузка файла
uploaded_file = st.file_uploader(
    "Выберите PDF файл маршрутного листа",
    type=['pdf'],
    help="Поддерживаются только PDF файлы"
)

# Информация о файле
if uploaded_file is not None:
    # Показываем информацию о файле
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Имя файла", uploaded_file.name[:20] + "..." if len(uploaded_file.name) > 20 else uploaded_file.name)
    
    with col2:
        file_size_mb = uploaded_file.size / (1024 * 1024)
        st.metric("Размер", f"{file_size_mb:.2f} MB")
    
    with col3:
        st.metric("Тип", "PDF")
    
    # Кнопка обработки
    st.markdown("---")
    
    if st.button("🚀 Начать парсинг PDF", type="primary", use_container_width=True):
        try:
            # Обрабатываем файл
            with st.spinner("Парсинг PDF файла... Это может занять несколько секунд"):
                excel_data, output_filename = process_pdf_file(uploaded_file)
            
            # Показываем успешное сообщение
            st.markdown('<div class="success-msg">✅ PDF файл успешно обработан и преобразован в Excel!</div>', unsafe_allow_html=True)
            
            # Кнопка скачивания
            st.download_button(
                label=f"⬇️ Скачать Excel файл: {output_filename}",
                data=excel_data,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            # Информация о завершении
            st.balloons()
            
            # Дополнительная информация
            st.info("💡 Файл содержит распарсенные данные из маршрутного листа в табличном формате.")
            
        except Exception as e:
            st.markdown(f'<div class="error-msg">❌ Ошибка обработки: {str(e)}</div>', unsafe_allow_html=True)
            
            # Кнопка для повторной попытки
            if st.button("🔄 Попробовать снова", type="secondary"):
                st.rerun()

# Боковая панель с информацией
with st.sidebar:
    st.header("ℹ️ О парсере")
    
    st.markdown("""
    ### Как это работает:
    1. **Загрузите PDF** маршрутного листа
    2. **Система автоматически** найдет таблицу
    3. **Извлечет данные** по сетке координат
    4. **Сохранит в Excel** с форматированием
    
    ### Технологии:
    - **PyMuPDF** для чтения PDF
    - **Pandas** для обработки данных
    - **OpenPyXL** для создания Excel
    
    ### Ограничения:
    - Только первая страница
    - Строгий формат заголовка
    - До 50MB на файл
    """)
    
    # Проверка наличия скрипта
    st.markdown("---")
    if os.path.exists("your_script.py"):
        st.success("✅ Скрипт парсера найден")
        try:
            # Проверяем импорт
            my_script = import_my_script()
            if my_script and hasattr(my_script, 'process'):
                st.success("✅ Функция process() доступна")
            else:
                st.error("❌ Функция process() не найдена")
        except:
            st.error("❌ Ошибка импорта скрипта")
    else:
        st.error("❌ Скрипт your_script.py не найден")
    
    # Время
    st.markdown("---")
    st.caption(f"Время: {datetime.now().strftime('%H:%M:%S')}")

# Футер
st.markdown("---")
st.caption("PDF Парсер маршрутных листов | Создано с Streamlit")

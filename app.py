import streamlit as st
import os
import tempfile
import sys
import importlib.util
import time
from datetime import datetime

# Настройки страницы
st.set_page_config(
    page_title="PDF Парсер авиационных листов",
    page_icon="✈️",
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
        margin-top: 20px;
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
        border-radius: 8px;
        border-left: 4px solid #007bff;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.title("✈️ Парсер авиационных листов")
st.markdown("Загрузите два PDF файла для автоматического парсинга в Excel")

# Информация о формате
st.markdown("""
<div class="info-box">
<h4>📌 Требования к файлам:</h4>
<ul>
<li><strong>Файл 1:</strong> Должен содержать слово "Takeoff" в начале (Takeoff file)</li>
<li><strong>Файл 2:</strong> Не должен содержать "Takeoff" (Main route file)</li>
<li>Оба файла должны быть в формате PDF</li>
<li>Будет обработана только первая страница каждого файла</li>
</ul>
</div>
""", unsafe_allow_html=True)

# Функция для импорта скрипта
def import_my_script():
    """Импортирует ваш скрипт обработки"""
    script_name = "your_script.py"
    
    if not os.path.exists(script_name):
        return None
    
    try:
        spec = importlib.util.spec_from_file_location("my_script", script_name)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    except Exception as e:
        st.error(f"Ошибка загрузки скрипта: {str(e)}")
        return None

# Основной интерфейс
st.markdown("---")
st.subheader("📤 Загрузка файлов")

# Загрузка двух файлов
col1, col2 = st.columns(2)

with col1:
    st.markdown("### Файл 1")
    st.markdown("*(должен содержать 'Takeoff')*")
    file1 = st.file_uploader(
        "Выберите первый PDF файл",
        type=['pdf'],
        key="file1",
        label_visibility="collapsed"
    )

with col2:
    st.markdown("### Файл 2")
    st.markdown("*(не должен содержать 'Takeoff')*")
    file2 = st.file_uploader(
        "Выберите второй PDF файл",
        type=['pdf'],
        key="file2",
        label_visibility="collapsed"
    )

# Отображение информации о загруженных файлах
if file1 is not None and file2 is not None:
    st.markdown("---")
    st.subheader("📋 Информация о файлах")
    
    # Проверка Takeoff для файла 1
    file1_preview = file1.getvalue()[:1000].decode('latin-1', errors='ignore')
    file1_has_takeoff = "takeoff" in file1_preview.lower()
    
    # Проверка Takeoff для файла 2
    file2_preview = file2.getvalue()[:1000].decode('latin-1', errors='ignore')
    file2_has_takeoff = "takeoff" in file2_preview.lower()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        <div class="file-info">
        <strong>{file1.name}</strong><br>
        Размер: {file1.size / 1024:.1f} KB<br>
        Takeoff: {'✅ Да' if file1_has_takeoff else '❌ Нет'}
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="file-info">
        <strong>{file2.name}</strong><br>
        Размер: {file2.size / 1024:.1f} KB<br>
        Takeoff: {'✅ Да' if file2_has_takeoff else '❌ Нет'}
        </div>
        """, unsafe_allow_html=True)
    
    # Проверка комбинации файлов
    if file1_has_takeoff and file2_has_takeoff:
        st.error("❌ Оба файла содержат 'Takeoff'. Нужен один файл с Takeoff и один без.")
    elif not file1_has_takeoff and not file2_has_takeoff:
        st.error("❌ Ни один файл не содержит 'Takeoff'. Нужен один файл с Takeoff.")
    else:
        st.success("✅ Файлы загружены правильно!")
        
        # Кнопка обработки
        if st.button("🚀 Начать обработку файлов", type="primary", use_container_width=True):
            try:
                # Импортируем скрипт
                my_script = import_my_script()
                
                if my_script is None:
                    raise Exception("Скрипт обработки не найден")
                
                if not hasattr(my_script, 'process'):
                    raise Exception("В скрипте не найдена функция 'process'")
                
                # Создаем прогресс-бар
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Обработка файлов
                with st.spinner("Обработка файлов..."):
                    status_text.text("🔄 Загружаю и анализирую файлы...")
                    progress_bar.progress(30)
                    
                    # Получаем содержимое файлов
                    file1_content = file1.getvalue()
                    file2_content = file2.getvalue()
                    
                    status_text.text("🔄 Выполняю парсинг PDF...")
                    progress_bar.progress(60)
                    
                    # Вызываем функцию обработки
                    excel_bytes = my_script.process(file1_content, file2_content)
                    
                    status_text.text("🔄 Форматирую результат...")
                    progress_bar.progress(90)
                    
                    # Генерируем имя выходного файла
                    if file1_has_takeoff:
                        base_name = os.path.splitext(file2.name)[0]
                    else:
                        base_name = os.path.splitext(file1.name)[0]
                    
                    output_filename = f"{base_name}_Flight_Log.xlsx"
                    
                    progress_bar.progress(100)
                    status_text.text("✅ Обработка завершена!")
                    time.sleep(0.5)
                
                # Очистка
                progress_bar.empty()
                status_text.empty()
                
                # Показываем успешное сообщение
                st.markdown('<div class="success-msg">✅ Файлы успешно обработаны и объединены в Excel!</div>', unsafe_allow_html=True)
                
                # Кнопка скачивания
                st.download_button(
                    label=f"⬇️ Скачать Excel файл: {output_filename}",
                    data=excel_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
                # Дополнительная информация
                st.info("""
                📊 Файл содержит:
                - Лист "Основное": ключевые данные из PDF
                - Лист "Main_Route_Grid": распарсенная таблица маршрута
                """)
                
                # Анимация успеха
                st.balloons()
                
            except Exception as e:
                st.markdown(f'<div class="error-msg">❌ Ошибка обработки: {str(e)}</div>', unsafe_allow_html=True)
                
                # Кнопка для повторной попытки
                if st.button("🔄 Попробовать снова", type="secondary"):
                    st.rerun()

elif file1 is not None or file2 is not None:
    st.warning("⚠️ Пожалуйста, загрузите оба файла для продолжения")

# Боковая панель
with st.sidebar:
    st.image("https://streamlit.io/images/brand/streamlit-mark-color.png", width=100)
    st.title("ℹ️ О сервисе")
    
    st.markdown("""
    ### Как это работает:
    1. **Загрузите 2 PDF файла**
    2. **Система определит** Takeoff файл
    3. **Обработает** основной файл
    4. **Создаст Excel** с результатами
    
    ### Особенности:
    - Автоматическое определение Takeoff
    - Парсинг таблиц по координатам
    - Два листа в Excel
    - Автоформатирование
    
    ### Технологии:
    - PyMuPDF для чтения PDF
    - Pandas для обработки данных
    - OpenPyXL для Excel
    """)
    
    # Проверка скрипта
    st.markdown("---")
    if os.path.exists("your_script.py"):
        try:
            my_script = import_my_script()
            if my_script and hasattr(my_script, 'process'):
                st.success("✅ Скрипт готов к работе")
            else:
                st.error("❌ Ошибка в скрипте")
        except:
            st.error("❌ Ошибка импорта скрипта")
    else:
        st.error("❌ Скрипт не найден")
    
    # Время
    st.markdown("---")
    st.caption(f"Время: {datetime.now().strftime('%H:%M:%S')}")

# Футер
st.markdown("---")
st.caption("✈️ Парсер авиационных листов | Создано с Streamlit")

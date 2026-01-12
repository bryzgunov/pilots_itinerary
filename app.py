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
    page_title="Обработчик файлов",
    page_icon="🔄",
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
</style>
""", unsafe_allow_html=True)

# Заголовок
st.title("📁 Обработчик файлов")
st.markdown("Загрузите файл, и он будет автоматически обработан")

# Функция для импорта вашего скрипта
def import_my_script():
    """Пытается импортировать ваш скрипт обработки"""
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

# Функция обработки файла
def process_uploaded_file(uploaded_file):
    """Обрабатывает загруженный файл"""
    
    # Создаем прогресс-бар
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Шаг 1: Сохранение файла
    status_text.text("🔄 Сохраняю файл...")
    progress_bar.progress(20)
    
    # Создаем временный файл для ввода
    with tempfile.NamedTemporaryFile(delete=False, 
                                   suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_input:
        tmp_input.write(uploaded_file.getvalue())
        input_path = tmp_input.name
    
    # Шаг 2: Подготовка выходного файла
    status_text.text("🔄 Подготавливаю обработку...")
    progress_bar.progress(40)
    
    # Генерируем имя выходного файла
    original_name = os.path.splitext(uploaded_file.name)[0]
    extension = os.path.splitext(uploaded_file.name)[1] or ".processed"
    output_filename = f"{original_name}_processed{extension}"
    output_path = os.path.join(tempfile.gettempdir(), output_filename)
    
    # Шаг 3: Импорт и выполнение вашего скрипта
    status_text.text("🔄 Загружаю скрипт обработки...")
    progress_bar.progress(60)
    
    try:
        # Пытаемся импортировать ваш скрипт
        my_script = import_my_script()
        
        if my_script is None:
            # Если скрипта нет, используем демо-обработку
            status_text.text("⚠️ Скрипт не найден, использую демо-режим...")
            
            # Простая демо-обработка
            if uploaded_file.type and 'text' in uploaded_file.type.lower():
                # Текстовый файл
                with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                # Простая обработка
                processed_content = content + "\n\n[Обработано в демо-режиме]"
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(processed_content)
            else:
                # Бинарный файл - просто копируем
                shutil.copy(input_path, output_path)
        else:
            # Используем ваш скрипт
            status_text.text("🔄 Выполняю обработку...")
            
            # Вариант 1: Если есть функция process()
            if hasattr(my_script, 'process'):
                my_script.process(input_path, output_path)
                
            # Вариант 2: Если есть функция main()
            elif hasattr(my_script, 'main'):
                # Сохраняем оригинальные аргументы
                old_argv = sys.argv.copy()
                sys.argv = ["your_script.py", input_path, output_path]
                my_script.main()
                sys.argv = old_argv
                
            # Вариант 3: Если нет нужных функций
            else:
                st.warning("В скрипте нет функций process() или main(). Использую демо-режим.")
                shutil.copy(input_path, output_path)
        
        progress_bar.progress(80)
        
        # Шаг 4: Проверка результата
        status_text.text("🔄 Проверяю результат...")
        
        if not os.path.exists(output_path):
            raise Exception("Обработанный файл не создан")
        
        progress_bar.progress(100)
        status_text.text("✅ Обработка завершена!")
        time.sleep(0.5)
        
        # Очистка
        progress_bar.empty()
        status_text.empty()
        
        # Читаем результат
        with open(output_path, 'rb') as f:
            processed_data = f.read()
        
        # Удаляем временные файлы
        try:
            os.unlink(input_path)
            os.unlink(output_path)
        except:
            pass
        
        return processed_data, output_filename
    
    except Exception as e:
        # Очистка при ошибке
        progress_bar.empty()
        status_text.empty()
        
        # Удаляем временные файлы
        try:
            if os.path.exists(input_path):
                os.unlink(input_path)
            if os.path.exists(output_path):
                os.unlink(output_path)
        except:
            pass
        
        raise e

# Основной интерфейс
st.markdown("---")

# Загрузка файла
uploaded_file = st.file_uploader(
    "Выберите файл для обработки",
    type=None,
    help="Поддерживаются любые типы файлов"
)

# Информация о файле
if uploaded_file is not None:
    # Показываем информацию о файле
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Имя файла", uploaded_file.name)
    
    with col2:
        file_size_mb = uploaded_file.size / (1024 * 1024)
        st.metric("Размер", f"{file_size_mb:.2f} MB")
    
    with col3:
        file_type = uploaded_file.type if uploaded_file.type else "Неизвестно"
        st.metric("Тип", file_type)
    
    # Кнопка обработки (или автоматическая обработка)
    st.markdown("---")
    
    # Автоматически начинаем обработку или ждем нажатия кнопки
    auto_process = st.checkbox("Начать обработку сразу после загрузки", value=True)
    
    if auto_process or st.button("🚀 Начать обработку", type="primary", use_container_width=True):
        try:
            # Обрабатываем файл
            with st.spinner("Обработка файла..."):
                processed_data, output_filename = process_uploaded_file(uploaded_file)
            
            # Показываем успешное сообщение
            st.markdown('<div class="success-msg">✅ Файл успешно обработан!</div>', unsafe_allow_html=True)
            
            # Кнопка скачивания
            st.download_button(
                label=f"⬇️ Скачать обработанный файл: {output_filename}",
                data=processed_data,
                file_name=output_filename,
                mime=uploaded_file.type or "application/octet-stream",
                type="primary",
                use_container_width=True
            )
            
            # Информация о завершении
            st.balloons()
            
        except Exception as e:
            st.markdown(f'<div class="error-msg">❌ Ошибка: {str(e)}</div>', unsafe_allow_html=True)
            
            # Кнопка для повторной попытки
            if st.button("🔄 Попробовать снова", type="secondary"):
                st.rerun()

# Боковая панель с информацией
with st.sidebar:
    st.header("ℹ️ Информация")
    
    st.markdown("""
    ### Как это работает:
    1. **Загрузите файл** с вашего устройства
    2. **Система автоматически** его обработает
    3. **Скачайте результат**
    
    ### Ваш скрипт:
    - Создайте файл `your_script.py`
    - Добавьте функцию `process(input_path, output_path)`
    - Загрузите в ту же папку
    
    ### Поддержка:
    - Все типы файлов
    - До 200MB на файл
    - Автоудаление после обработки
    """)
    
    # Проверка наличия скрипта
    st.markdown("---")
    if os.path.exists("your_script.py"):
        st.success("✅ Скрипт your_script.py найден")
    else:
        st.warning("⚠️ Скрипт your_script.py не найден")
        st.info("Создайте файл your_script.py с функцией process()")
    
    # Время
    st.markdown("---")
    st.caption(f"Время: {datetime.now().strftime('%H:%M:%S')}")

# Футер
st.markdown("---")
st.caption("Веб-сервис для обработки файлов | Создано с Streamlit")

import streamlit as st
import os
import tempfile
import sys
import importlib.util
import shutil
from datetime import datetime

# Импорт вашего скрипта
def import_script(script_path):
    """Динамический импорт Python-скрипта"""
    spec = importlib.util.spec_from_file_location("custom_script", script_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

# Конфигурация страницы
st.set_page_config(
    page_title="Обработчик файлов",
    page_icon="🔄",
    layout="wide"
)

# Стили
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton button {
        width: 100%;
        height: 3rem;
        font-size: 1.2rem;
    }
    .success-box {
        padding: 20px;
        background-color: #d4edda;
        border-radius: 10px;
        border: 1px solid #c3e6cb;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.markdown('<h1 class="main-header">🔄 Обработчик файлов</h1>', unsafe_allow_html=True)
st.markdown("---")

# Боковая панель
with st.sidebar:
    st.image("https://streamlit.io/images/brand/streamlit-mark-color.png", width=100)
    st.title("ℹ️ О сервисе")
    
    st.markdown("""
    ### 📝 Описание
    Этот сервис обрабатывает ваши файлы
    с помощью кастомного Python-скрипта.
    
    ### ⚙️ Технологии
    - **Frontend**: Streamlit
    - **Хостинг**: Streamlit Cloud
    - **Обработка**: Python
    
    ### 🛡️ Безопасность
    - Файлы удаляются после обработки
    - Нет постоянного хранения
    - Все операции временные
    """)
    
    # Статус
    st.divider()
    st.caption(f"🕐 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.caption("v1.0 | Streamlit Cloud")

# Основной интерфейс
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    st.info("""
    ### 📱 Как использовать:
    1. Нажмите кнопку ниже
    2. Загрузите файл с телефона/компьютера
    3. Дождитесь обработки
    4. Скачайте результат
    """)

# Кнопка запуска
if st.button("🚀 НАЧАТЬ ОБРАБОТКУ ФАЙЛА", type="primary", use_container_width=True):
    
    # Загрузка файла
    uploaded_file = st.file_uploader(
        "📤 Перетащите файл или нажмите для выбора",
        type=None,
        help="Максимальный размер файла: 200MB",
        key="file_uploader_main"
    )
    
    if uploaded_file is not None:
        # Информация о файле
        file_details = {
            "Имя файла": uploaded_file.name,
            "Тип файла": uploaded_file.type or "Не определен",
            "Размер": f"{uploaded_file.size / 1024:.2f} KB"
        }
        
        st.write("📋 **Информация о файле:**")
        st.json(file_details)
        
        # Кнопка обработки
        if st.button("🔄 ОБРАБОТАТЬ ФАЙЛ", type="secondary", use_container_width=True):
            with st.spinner("⏳ Идет обработка файла... Пожалуйста, подождите"):
                # Прогресс-бар
                progress_bar = st.progress(0)
                
                # Создаем временные файлы
                import tempfile
                
                # Входной файл
                with tempfile.NamedTemporaryFile(delete=False, 
                                               suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    input_path = tmp_in.name
                
                progress_bar.progress(30)
                
                # Выходной файл
                original_name = os.path.splitext(uploaded_file.name)[0]
                extension = os.path.splitext(uploaded_file.name)[1] or ".processed"
                output_filename = f"{original_name}_ОБРАБОТАННЫЙ{extension}"
                output_path = os.path.join(tempfile.gettempdir(), output_filename)
                
                try:
                    progress_bar.progress(50)
                    
                    # ВАША ЛОГИКА ОБРАБОТКИ - ВАРИАНТ 1: Использование вашего скрипта
                    if os.path.exists("your_script.py"):
                        st.info("🔍 Обнаружен ваш скрипт: your_script.py")
                        
                        # Импортируем ваш скрипт
                        try:
                            your_module = import_script("your_script.py")
                            
                            # Вызываем функцию (адаптируйте под ваш случай)
                            if hasattr(your_module, 'process'):
                                your_module.process(input_path, output_path)
                                st.success("✅ Обработка через функцию process() завершена")
                                
                            elif hasattr(your_module, 'main'):
                                # Сохраняем старые аргументы
                                old_argv = sys.argv
                                sys.argv = ["your_script.py", input_path, output_path]
                                your_module.main()
                                sys.argv = old_argv
                                st.success("✅ Обработка через main() завершена")
                                
                            else:
                                st.warning("⚠️ Не найдены функции process() или main(), копируем файл")
                                shutil.copy(input_path, output_path)
                                
                        except Exception as script_error:
                            st.error(f"❌ Ошибка в вашем скрипте: {str(script_error)}")
                            # Резервный вариант - копирование
                            shutil.copy(input_path, output_path)
                    
                    # ВАРИАНТ 2: Если скрипта нет - демо-обработка
                    else:
                        st.warning("⚠️ Файл your_script.py не найден, выполняем демо-обработку")
                        
                        # Определяем тип файла
                        if uploaded_file.type and 'text' in uploaded_file.type:
                            # Текстовый файл
                            with open(input_path, 'r', encoding='utf-8', errors='ignore') as f_in:
                                content = f_in.read()
                            
                            # Простая обработка текста
                            processed_content = f"{content}\n\n[Processed via Streamlit Cloud]"
                            
                            with open(output_path, 'w', encoding='utf-8') as f_out:
                                f_out.write(processed_content)
                                
                        else:
                            # Бинарный файл - просто копируем
                            shutil.copy(input_path, output_path)
                    
                    progress_bar.progress(80)
                    
                    # Проверяем, создан ли выходной файл
                    if not os.path.exists(output_path):
                        raise FileNotFoundError("Выходной файл не создан")
                    
                    # Узнаем размер выходного файла
                    output_size = os.path.getsize(output_path)
                    
                    progress_bar.progress(100)
                    
                    # Успешное завершение
                    st.markdown('<div class="success-box">', unsafe_allow_html=True)
                    st.success("✅ Файл успешно обработан!")
                    st.markdown(f"**Размер обработанного файла:** {output_size / 1024:.2f} KB")
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # Кнопка скачивания
                    with open(output_path, 'rb') as f:
                        file_bytes = f.read()
                    
                    st.download_button(
                        label=f"⬇️ СКАЧАТЬ: {output_filename}",
                        data=file_bytes,
                        file_name=output_filename,
                        mime=uploaded_file.type or "application/octet-stream",
                        type="primary",
                        use_container_width=True
                    )
                    
                    # Очистка временных файлов
                    try:
                        os.unlink(input_path)
                        os.unlink(output_path)
                        st.caption("🗑️ Временные файлы удалены")
                    except:
                        pass
                    
                except Exception as e:
                    progress_bar.progress(100)
                    st.error(f"❌ Ошибка при обработке файла")
                    
                    # Детальная информация об ошибке
                    with st.expander("🔧 Детали ошибки"):
                        st.code(f"""
Ошибка: {str(e)}
Тип: {type(e).__name__}

Путь к входному файлу: {input_path}
Путь к выходному файлу: {output_path}

Размер входного файла: {uploaded_file.size} байт
Тип файла: {uploaded_file.type}
                        """)
                    
                    # Пытаемся очистить временные файлы
                    try:
                        if os.path.exists(input_path):
                            os.unlink(input_path)
                        if os.path.exists(output_path):
                            os.unlink(output_path)
                    except:
                        pass

# Разделитель
st.markdown("---")

# Инструкция по настройке
with st.expander("⚙️ Как подключить свой скрипт?"):
    st.markdown("""
    ### 1. Создайте файл `your_script.py`
    
    ```python
    # your_script.py
    import sys
    
    def process(input_path, output_path):
        \"""
        Обработка файла
        \"""
        # Ваш код обработки здесь
        with open(input_path, 'r') as f_in:
            data = f_in.read()
        
        # Пример: преобразование текста
        processed = data.upper()
        
        with open(output_path, 'w') as f_out:
            f_out.write(processed)
    
    # ИЛИ если используете main()
    def main():
        if len(sys.argv) == 3:
            process(sys.argv[1], sys.argv[2])
        else:
            print("Использование: python your_script.py input output")
    
    if __name__ == "__main__":
        main()
    ```
    
    ### 2. Добавьте зависимости в `requirements.txt`
    
    ```txt
    streamlit>=1.28.0
    # ваши библиотеки
    pandas>=2.0.0
    numpy>=1.24.0
    ```
    
    ### 3. Загрузите оба файла в репозиторий GitHub
    """)

# Футер
st.markdown("---")
st.caption("✨ Веб-сервис для обработки файлов | Создано с помощью Streamlit")
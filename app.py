import streamlit as st
import os
import tempfile
import sys
import importlib.util

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
</style>
""", unsafe_allow_html=True)

# Заголовок
st.markdown('<h1 class="main-header">🔄 Обработчик файлов</h1>', unsafe_allow_html=True)
st.markdown("---")

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
        help="Максимальный размер файла: 200MB"
    )
    
    if uploaded_file is not None:
        # Информация о файле
        file_details = {
            "Имя файла": uploaded_file.name,
            "Тип файла": uploaded_file.type,
            "Размер": f"{uploaded_file.size / 1024:.2f} KB"
        }
        
        st.json(file_details)
        
        with st.spinner("⏳ Идет обработка файла... Пожалуйста, подождите"):
            
            # Создаем временные файлы
            import tempfile
            
            # Входной файл
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_in:
                tmp_in.write(uploaded_file.getvalue())
                input_path = tmp_in.name
            
            # Выходной файл
            original_name = os.path.splitext(uploaded_file.name)[0]
            extension = os.path.splitext(uploaded_file.name)[1] or ".processed"
            output_filename = f"{original_name}_ОБРАБОТАННЫЙ{extension}"
            output_path = os.path.join(tempfile.gettempdir(), output_filename)
            
            try:
                # ЗДЕСЬ ВАША ЛОГИКА ОБРАБОТКИ
                # Пример 1: Если ваш скрипт - отдельный файл
                if os.path.exists("your_script.py"):
                    # Импортируем ваш скрипт
                    your_module = import_script("your_script.py")
                    
                    # Вызываем функцию (адаптируйте под ваш случай)
                    if hasattr(your_module, 'process'):
                        your_module.process(input_path, output_path)
                    elif hasattr(your_module, 'main'):
                        your_module.main(input_path, output_path)
                    else:
                        # Если скрипт не имеет функций, копируем как есть
                        import shutil
                        shutil.copy(input_path, output_path)
                
                # Пример 2: Если обработка простая
                else:
                    with open(input_path, 'rb') as f_in, open(output_path, 'wb') as f_out:
                        # Ваша логика обработки
                        data = f_in.read()
                        # Например, просто добавляем метку
                        processed_data = data + b"\n\n[Обработано через Streamlit Cloud]"
                        f_out.write(processed_data)
                
                # Успех
                st.success("✅ Файл успешно обработан!")
                
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
                
                # Очистка
                os.unlink(input_path)
                os.unlink(output_path)
                
            except Exception as e:
                st.error(f"❌ Ошибка при обработке: {str(e)}")
                st.code(str(e), language="python")

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
    from datetime import datetime
    st.divider()
    st.caption(f"🕐 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.caption("v1.0 | Streamlit Cloud")
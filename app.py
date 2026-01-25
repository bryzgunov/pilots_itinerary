import streamlit as st
import os
import tempfile
from your_script import process_two_pdfs
from datetime import datetime
import time

# Настройки страницы
st.set_page_config(
    page_title="Advanced Flight Log Processor",
    page_icon="✈️",
    layout="wide"
)

# Стили
st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1rem;
    }
    .subtitle {
        font-size: 1.2rem;
        color: #6B7280;
        text-align: center;
        margin-bottom: 2rem;
    }
    .file-card {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        border: 2px dashed #dee2e6;
        margin: 10px 0;
    }
    .success-card {
        background-color: #d4edda;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #c3e6cb;
    }
    .error-card {
        background-color: #f8d7da;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #f5c6cb;
    }
    .info-card {
        background-color: #d1ecf1;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #bee5eb;
    }
    .stButton > button {
        font-size: 1.1rem;
        padding: 10px 20px;
    }
    .sheet-badge {
        display: inline-block;
        background-color: #6c757d;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.8rem;
        margin: 2px;
    }
    .progress-container {
        margin: 20px 0;
        padding: 15px;
        background-color: #f8f9fa;
        border-radius: 10px;
        border: 1px solid #dee2e6;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.markdown('<h1 class="main-title">✈️ Advanced Flight Log Processor</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Upload two PDF files to generate a comprehensive 6-sheet flight log report</p>', unsafe_allow_html=True)

# Информация о системе
st.markdown("""
<div class="info-card">
<h4>📋 What this advanced tool does:</h4>
<ul>
<li><b>1. Takes two PDF files</b> - one with Takeoff data and one with main route</li>
<li><b>2. Automatically detects</b> which file contains Takeoff information</li>
<li><b>3. Creates a comprehensive Excel report</b> with <b>6 sheets</b>:</li>
<div style="margin-left: 20px;">
<div><span class="sheet-badge">Основное</span> - Basic flight information</div>
<div><span class="sheet-badge">Main_Route_Grid</span> - Parsed route table</div>
<div><span class="sheet-badge">Airport_Table</span> - Airport information table</div>
<div><span class="sheet-badge">Airport_Maps</span> - Airport diagrams and maps</div>
<div><span class="sheet-badge">ForeFlight</span> - Takeoff data analysis</div>
<div><span class="sheet-badge">Generated_Sheet</span> - Formatted flight log with offsets</div>
</div>
</ul>
</div>
""", unsafe_allow_html=True)

# Загрузка файлов
st.markdown("---")
st.subheader("📤 Upload PDF Files")

col1, col2 = st.columns(2)

with col1:
    uploaded_file1 = st.file_uploader(
        "First PDF file",
        type=['pdf'],
        help="PDF file (either Takeoff or main route)",
        key="file1"
    )

with col2:
    uploaded_file2 = st.file_uploader(
        "Second PDF file", 
        type=['pdf'],
        help="PDF file (the other one of the pair)",
        key="file2"
    )

# Отображение информации о файлах
if uploaded_file1 and uploaded_file2:
    st.markdown("---")
    st.subheader("📋 Uploaded Files")
    
    # Создаем карточки для файлов
    file_col1, file_col2 = st.columns(2)
    
    with file_col1:
        st.markdown(f"""
        <div class="file-card">
        <h4>📄 File 1</h4>
        <p><b>Name:</b> {uploaded_file1.name}</p>
        <p><b>Size:</b> {uploaded_file1.size / 1024:.1f} KB</p>
        <p><b>Type:</b> PDF</p>
        </div>
        """, unsafe_allow_html=True)
    
    with file_col2:
        st.markdown(f"""
        <div class="file-card">
        <h4>📄 File 2</h4>
        <p><b>Name:</b> {uploaded_file2.name}</p>
        <p><b>Size:</b> {uploaded_file2.size / 1024:.1f} KB</p>
        <p><b>Type:</b> PDF</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Проверка на одинаковые имена
    if uploaded_file1.name == uploaded_file2.name:
        st.error("❌ Error: Files have the same name. Please upload different files.")
    else:
        # Дополнительная информация
        st.info("💡 The system will automatically detect which file contains 'Takeoff' information.")
        
        # Кнопка обработки
        st.markdown("---")
        st.subheader("🚀 Processing")
        
        if st.button("Start Advanced Processing", type="primary", use_container_width=True):
            try:
                # Контейнер для прогресса
                progress_container = st.container()
                
                with progress_container:
                    st.markdown("### Processing Progress")
                    
                    # Показываем прогресс
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Шаг 1: Анализ файлов
                    status_text.text("Step 1/6: Analyzing files and detecting Takeoff...")
                    progress_bar.progress(15)
                    time.sleep(0.5)
                    
                    # Шаг 2: Чтение PDF
                    status_text.text("Step 2/6: Reading and parsing PDF files...")
                    progress_bar.progress(30)
                    
                    # Получаем байты файлов
                    file1_bytes = uploaded_file1.getvalue()
                    file2_bytes = uploaded_file2.getvalue()
                    
                    # Шаг 3: Обработка
                    status_text.text("Step 3/6: Processing main route data...")
                    progress_bar.progress(45)
                    
                    # Обрабатываем файлы
                    excel_bytes = process_two_pdfs(
                        file1_bytes, 
                        file2_bytes,
                        uploaded_file1.name,
                        uploaded_file2.name
                    )
                    
                    # Шаг 4: Создание дополнительных листов
                    status_text.text("Step 4/6: Creating Airport and Takeoff sheets...")
                    progress_bar.progress(60)
                    time.sleep(0.5)
                    
                    # Шаг 5: Генерация итогового листа
                    status_text.text("Step 5/6: Generating final formatted sheet...")
                    progress_bar.progress(75)
                    time.sleep(0.5)
                    
                    # Шаг 6: Завершение
                    status_text.text("Step 6/6: Finalizing and preparing download...")
                    progress_bar.progress(100)
                    time.sleep(0.5)
                
                # Успешное завершение
                st.markdown('<div class="success-card">', unsafe_allow_html=True)
                st.success("✅ Advanced processing completed successfully!")
                st.markdown("</div>", unsafe_allow_html=True)
                
                # Генерируем имя выходного файла
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_filename = f"Flight_Log_Report_Advanced_{timestamp}.xlsx"
                
                # Информация о созданном файле
                st.markdown("""
                <div class="info-card">
                <h4>📊 Generated Advanced Report Contains 6 Sheets:</h4>
                <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px;">
                <div><span class="sheet-badge">Основное</span> Basic flight info</div>
                <div><span class="sheet-badge">Main_Route_Grid</span> Route table</div>
                <div><span class="sheet-badge">Airport_Table</span> Airport data</div>
                <div><span class="sheet-badge">Airport_Maps</span> Airport diagrams</div>
                <div><span class="sheet-badge">ForeFlight</span> Takeoff analysis</div>
                <div><span class="sheet-badge">Generated_Sheet</span> Formatted log</div>
                </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Кнопка скачивания
                st.download_button(
                    label=f"⬇️ Download Advanced Excel Report: {output_filename}",
                    data=excel_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
                # Дополнительная информация
                st.info("""
                **Advanced Features:**
                - **Takeoff Data Analysis**: Extracts and processes Takeoff performance data
                - **Airport Maps**: Extracts airport diagrams from PDF
                - **Formatted Flight Log**: Creates professional flight log with proper formatting
                - **Data Integration**: Combines data from both PDF files intelligently
                """)
                
                # Анимация успеха
                st.balloons()
                
            except Exception as e:
                st.markdown('<div class="error-card">', unsafe_allow_html=True)
                st.error(f"❌ Processing Error: {str(e)}")
                st.markdown("</div>", unsafe_allow_html=True)
                
                # Дополнительная информация об ошибке
                st.warning("""
                **Troubleshooting tips:**
                1. Ensure both PDF files are valid and not corrupted
                2. Make sure one file contains 'Takeoff' information
                3. Check that files are not password protected
                4. Try with smaller files if possible
                """)
                
                # Кнопка для повторной попытки
                if st.button("🔄 Try Again", type="secondary"):
                    st.rerun()

# Боковая панель
with st.sidebar:
    st.header("ℹ️ About Advanced Version")
    
    st.markdown("""
    ### ✈️ Advanced Flight Log Processor
    This advanced tool processes flight log PDF files and creates comprehensive Excel reports with 6 sheets.
    
    ### 📁 Input Requirements:
    - **Two PDF files** (one with Takeoff, one with main route)
    - **PDF format** from flight planning systems
    - **Maximum size**: 50MB per file
    
    ### 📊 Output Sheets:
    1. **Основное** - Basic flight info
    2. **Main_Route_Grid** - Route table
    3. **Airport_Table** - Airport data
    4. **Airport_Maps** - Airport diagrams
    5. **ForeFlight** - Takeoff analysis
    6. **Generated_Sheet** - Formatted log
    
    ### ⚙️ Technology Stack:
    - **PyMuPDF** - Advanced PDF parsing
    - **Pandas** - Data processing
    - **OpenPyXL** - Excel generation with images
    - **Pillow** - Image processing
    - **Streamlit** - Web interface
    
    ### 🔒 Privacy & Security:
    - Files processed in memory only
    - No permanent storage
    - All data deleted after processing
    - Secure HTTPS connection
    """)
    
    # Проверка скрипта
    st.markdown("---")
    if os.path.exists("your_script.py"):
        file_size = os.path.getsize("your_script.py") / 1024
        st.success(f"✅ Advanced script loaded ({file_size:.1f} KB)")
        st.info(f"Version: 6-sheet advanced processor")
    else:
        st.error("❌ Script not found")
    
    # Информация о версии
    st.markdown("---")
    st.caption(f"Version: 4.0 (Advanced 6-sheet)")
    st.caption(f"Time: {datetime.now().strftime('%H:%M:%S')}")
    st.caption("Built for professional flight operations")

# Футер
st.markdown("---")
st.caption("✈️ Advanced Flight Log Processor | Professional aviation document processing | Created with Streamlit")

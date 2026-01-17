import fitz  # PyMuPDF
import unicodedata
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import io

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
def normalize_ascii(text):
    """Нормализация текста для проверки на ASCII"""
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in nfkd if ord(c) < 128)

def is_takeoff_file(pdf_bytes):
    """Определяет, содержит ли файл 'Takeoff' в начале"""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        raw = doc[0].get_text("text")[:250]
        doc.close()
        return normalize_ascii(raw).strip().lower().startswith("takeoff")
    except Exception as e:
        print(f"Ошибка при проверке файла на Takeoff: {e}")
        return False

def extract_first_n_lines_from_doc(doc, n=32):
    """Извлекает первые n строк из PDF документа"""
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
    lines = []
    for block in blocks:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            text = "".join(span["text"] for span in line["spans"]).strip()
            if text:
                lines.append(text)
                if len(lines) >= n:
                    return lines
    return lines

def parse_main_route_table(doc):
    """Парсит основную таблицу маршрута из PDF"""
    page = doc[0]
    all_words = page.get_text("words")

    # --- Поиск заголовка ---
    target_y = None
    for word_tuple in all_words:
        x0, y0, x1, y1, text, *_ = word_tuple
        if text == "WAYPOINT":
            for w in all_words:
                wx0, wy0, wx1, wy1, wtext, *_ = w
                if wtext == "ACT" and abs((y0 + y1)/2 - (wy0 + wy1)/2) < 5 and wx0 > x0:
                    target_y = (y0 + y1) / 2
                    break
            if target_y is not None:
                break

    if target_y is None:
        for word_tuple in all_words:
            x0, y0, x1, y1, text, *_ = word_tuple
            if text == "MAG":
                target_y = (y0 + y1) / 2 + 15
                break

    if target_y is None:
        raise RuntimeError("Не найдена строка заголовка таблицы.")

    # --- Извлечение заголовков ---
    header_keywords = ["WAYPOINT", "AIRWAY", "HDG", "CRS", "ALT", "CMP", "DIR/SPD", "ISA", "TAS", "GS", "LEG", "REM", "USED", "ACT", "ETE"]
    header_words_info = []
    tolerance = 5.0
    for word_tuple in all_words:
        x0, y0, x1, y1, text, *_ = word_tuple
        center_y = (y0 + y1) / 2
        if abs(center_y - target_y) <= tolerance and text in header_keywords:
            header_words_info.append((text, x0, x1))

    header_words_info.sort(key=lambda item: item[1])

    # --- Построение XX ---
    XX = []
    for i in range(1, len(header_words_info)):
        x1_prev = header_words_info[i-1][2]
        x0_next = header_words_info[i][1]
        boundary_x = (x0_next - x1_prev) / 2 + x1_prev
        XX.append(boundary_x)

    if XX:
        x0_airway = next((x0 for text, x0, x1 in header_words_info if text == "AIRWAY"), None)
        if x0_airway is not None:
            XX[0] = x0_airway - 2
        XX.insert(0, 5)
        XX.append(XX[-1] + 10)

    # --- Построение YY ---
    alt_coords = None
    for text, x0, x1 in header_words_info:
        if text == "ALT":
            for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
                if wtext == "ALT" and abs(wx0 - x0) < 1 and abs(wx1 - x1) < 1:
                    alt_coords = (wx0, wy0, wx1, wy1)
                    break
            if alt_coords:
                break

    if not alt_coords:
        raise RuntimeError("Не найдены координаты слова 'ALT'.")

    x0_alt, y0_alt, x1_alt, y1_alt = alt_coords

    y0_alternate = None
    for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
        if "ALTERNATE" in wtext:
            y0_alternate = wy0
            break
    if y0_alternate is None:
        for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
            if "2000 FT" in wtext and "ISA:" in wtext:
                y0_alternate = wy0
                break

    if y0_alternate is None:
        raise RuntimeError("Не найдена нижняя граница таблицы.")

    YY = []
    for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
        if x0_alt <= (wx0 + wx1) / 2 <= x1_alt and y1_alt <= wy0 <= y0_alternate:
            if wtext != "ALT" and "ALTERNATE" not in wtext and "2000 FT" not in wtext:
                YY.append(wy0 - 2)
    YY.append(y0_alternate - 2)

    # --- Парсинг сетки ---
    num_cols = len(XX) - 1
    num_rows = len(YY) - 1

    exact_columns = [
        "WAYPOINT", "AIRWAY", "HDG", "CRS", "ALT", "CMP", "DIR/SPD", "ISA",
        "TAS", "GS", "LEG", "REM", "USED", "REM", "ACT", "LEG", "REM", "ETE", "ACT"
    ]

    data_grid = []
    if num_cols <= len(exact_columns):
        for row_idx in range(num_rows):
            row_data = [''] * len(exact_columns)
            for col_idx in range(min(num_cols, len(exact_columns))):
                x_min = XX[col_idx]
                x_max = XX[col_idx + 1]
                y_min = YY[row_idx]
                y_max = YY[row_idx + 1]
                cell_texts = []
                for word_tuple in all_words:
                    wx0, wy0, wx1, wy1, wtext, *_ = word_tuple
                    center_x = (wx0 + wx1) / 2
                    center_y = (wy0 + wy1) / 2
                    if x_min <= center_x <= x_max and y_min <= center_y <= y_max:
                        cell_texts.append(wtext)
                row_data[col_idx] = ' '.join(cell_texts) if cell_texts else ''
            data_grid.append(row_data)
    else:
        for row_idx in range(num_rows):
            row_data = []
            for col_idx in range(len(exact_columns)):
                x_min = XX[col_idx]
                x_max = XX[col_idx + 1]
                y_min = YY[row_idx]
                y_max = YY[row_idx + 1]
                cell_texts = []
                for word_tuple in all_words:
                    wx0, wy0, wx1, wy1, wtext, *_ = word_tuple
                    center_x = (wx0 + wx1) / 2
                    center_y = (wy0 + wy1) / 2
                    if x_min <= center_x <= x_max and y_min <= center_y <= y_max:
                        cell_texts.append(wtext)
                row_data.append(' '.join(cell_texts) if cell_texts else '')
            data_grid.append(row_data)

    df = pd.DataFrame(data_grid, columns=exact_columns)
    return df

def process_two_pdfs(file1_bytes, file2_bytes, file1_name, file2_name):
    """
    Основная функция обработки двух PDF файлов
    Вход: два PDF файла в виде байтов
    Выход: Excel файл в виде байтов
    """
    print(f"Начинаю обработку файлов: {file1_name} и {file2_name}")
    
    # Определяем, какой файл является Takeoff
    file1_is_takeoff = is_takeoff_file(file1_bytes)
    file2_is_takeoff = is_takeoff_file(file2_bytes)
    
    print(f"Файл 1 ({file1_name}) содержит Takeoff: {file1_is_takeoff}")
    print(f"Файл 2 ({file2_name}) содержит Takeoff: {file2_is_takeoff}")
    
    # Проверяем, что ровно один файл содержит Takeoff
    if file1_is_takeoff and file2_is_takeoff:
        raise ValueError("Оба файла содержат 'Takeoff'. Нужен только один файл с Takeoff.")
    elif not file1_is_takeoff and not file2_is_takeoff:
        raise ValueError("Ни один из файлов не содержит 'Takeoff'. Нужен один файл с Takeoff.")
    
    # Определяем, какой файл обрабатывать (не Takeoff)
    if file1_is_takeoff:
        processing_file_bytes = file2_bytes
        processing_file_name = file2_name
        takeoff_file_name = file1_name
    else:
        processing_file_bytes = file1_bytes
        processing_file_name = file1_name
        takeoff_file_name = file2_name
    
    print(f"Обрабатываю файл: {processing_file_name}")
    print(f"Takeoff файл: {takeoff_file_name}")
    
    # Открываем PDF для обработки
    doc = fitz.open(stream=processing_file_bytes, filetype="pdf")
    
    # === ЛИСТ 1: ОСНОВНОЕ ===
    print("Создаю лист 'Основное'...")
    lines = extract_first_n_lines_from_doc(doc, n=32)
    while len(lines) < 32:
        lines.append("")

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Основное"

    ws1.cell(row=1, column=1, value=lines[0])
    ws1.cell(row=2, column=1, value=lines[1])
    ws1.cell(row=1, column=7, value=lines[2])
    ws1.cell(row=2, column=7, value=lines[3])

    block1 = lines[4:18]
    if len(block1) == 14:
        for col in range(7):
            ws1.cell(row=4, column=1 + col, value=block1[col * 2])
            ws1.cell(row=5, column=1 + col, value=block1[col * 2 + 1])

    block2 = lines[18:32]
    if len(block2) == 14:
        for col in range(7):
            ws1.cell(row=7, column=1 + col, value=block2[col * 2])
            ws1.cell(row=8, column=1 + col, value=block2[col * 2 + 1])

    bold_font = Font(bold=True)
    left_align = Alignment(horizontal="left", vertical="top")
    right_align = Alignment(horizontal="right", vertical="top")

    ws1['A1'].font = bold_font
    ws1['A1'].alignment = left_align
    ws1['G1'].alignment = right_align
    ws1['G2'].alignment = right_align

    for col in range(1, 8):
        ws1.cell(row=4, column=col).font = bold_font
        ws1.cell(row=4, column=col).alignment = left_align
        ws1.cell(row=7, column=col).font = bold_font
        ws1.cell(row=7, column=col).alignment = left_align

    for row in [2, 5, 8]:
        for col in range(1, 8):
            cell = ws1.cell(row=row, column=col)
            if cell.value is not None:
                cell.alignment = left_align

    col_widths = [12, 11, 20, 14, 15, 10, 13]
    for i, w in enumerate(col_widths, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    ws1.page_setup.orientation = 'portrait'
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    ws1.page_margins.left = 0.2
    ws1.page_margins.right = 0.2
    ws1.page_margins.top = 0.3
    ws1.page_margins.bottom = 0.3
    ws1.print_area = 'A1:G8'
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = False

    # === ЛИСТ 2: ПАРСИНГ ТАБЛИЦЫ ===
    print("Парсинг таблицы маршрута...")
    df = parse_main_route_table(doc)

    # === СОЗДАНИЕ ВТОРОГО ЛИСТА ===
    ws2 = wb.create_sheet(title="Main_Route_Grid")

    # Стили
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # Заголовки — строка 2
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws2.cell(row=2, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    # Данные — начиная со строки 3
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=3):
        for c_idx, value in enumerate(row, start=1):
            ws2.cell(row=r_idx, column=c_idx, value=value)

    # === СТИЛИЗАЦИЯ И ОБЪЕДИНЕНИЕ СТРОКИ 1 ===
    num_cols = len(df.columns)
    for col_idx in range(1, num_cols + 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    # Объединение и значения
    ws2.merge_cells(start_row=1, start_column=3, end_row=1, end_column=4)
    ws2.cell(row=1, column=3, value="MAG")

    ws2.merge_cells(start_row=1, start_column=6, end_row=1, end_column=7)
    ws2.cell(row=1, column=6, value="WIND")

    ws2.merge_cells(start_row=1, start_column=9, end_row=1, end_column=10)
    ws2.cell(row=1, column=9, value="SPD KT")

    ws2.merge_cells(start_row=1, start_column=11, end_row=1, end_column=12)
    ws2.cell(row=1, column=11, value="DIST NM")

    ws2.merge_cells(start_row=1, start_column=13, end_row=1, end_column=14)
    ws2.cell(row=1, column=13, value="FUEL G")

    ws2.merge_cells(start_row=1, start_column=16, end_row=1, end_column=18)
    ws2.cell(row=1, column=16, value="TIME")

    # === АВТОШИРИНА (без MergedCell ошибки) ===
    max_col = ws2.max_column
    for col_idx in range(1, max_col + 1):
        max_len = 0
        for row in ws2.iter_rows(min_col=col_idx, max_col=col_idx, min_row=1, max_row=ws2.max_row):
            cell = row[0]
            if hasattr(cell, 'value') and cell.value is not None:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        adjusted_width = min(max_len + 2, 50)
        ws2.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

    doc.close()

    # Сохраняем workbook в bytes
    excel_bytes = io.BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)
    
    print(f"✅ Обработка завершена! Создан Excel файл.")
    return excel_bytes.getvalue()

def process(input_path, output_path):
    """
    Функция для обратной совместимости (если нужна)
    Принимает один файл, но в нашем случае не используется
    """
    raise NotImplementedError("Для обработки двух файлов используйте process_two_pdfs")

def main():
    """Для локального тестирования"""
    import sys
    if len(sys.argv) != 3:
        print("Использование: python your_script.py <файл1.pdf> <файл2.pdf>")
        sys.exit(1)
    
    with open(sys.argv[1], 'rb') as f1, open(sys.argv[2], 'rb') as f2:
        result = process_two_pdfs(f1.read(), f2.read(), sys.argv[1], sys.argv[2])
        
    with open("output.xlsx", 'wb') as f:
        f.write(result)
    print("Файл сохранен как output.xlsx")

if __name__ == "__main__":
    main()

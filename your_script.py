import fitz  # PyMuPDF
import unicodedata
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins
from PIL import Image as PILImage
import io
import os
import numpy as np

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

def parse_airport_table(doc):
    """Парсит таблицу аэропортов"""
    # Ищем страницу с "AIRPORT"
    target_text = "AIRPORT"
    page_with_table = None
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        if target_text in text:
            page_with_table = page
            print(f"Ключевое слово '{target_text}' найдено на странице {page_num + 1}")
            break

    if page_with_table is None:
        print(f"Ключевое слово '{target_text}' не найдено.")
        return pd.DataFrame()  # Возвращаем пустой DataFrame

    # Поиск координат "AIRPORT"
    words = page_with_table.get_text("words")
    airport_coords = None
    for word in words:
        if word[4].upper() == target_text:
            airport_coords = (word[0], word[1], word[2], word[3])
            print(f"Координаты 'AIRPORT': x0={airport_coords[0]:.2f}, y0={airport_coords[1]:.2f}")
            break

    if airport_coords is None:
        print("Не удалось получить координаты слова 'AIRPORT'.")
        return pd.DataFrame()

    # Определяем XX
    XX_airport = [5, 75, 150, 200, 250, 325, 375, 425, 475, 525, 600]
    print("Массив XX (аэропорты):", XX_airport)

    # Определяем YY
    YY_airport = []
    YY_airport.append(airport_coords[3] + 2)

    dest_coords = None
    for word in words:
        if word[4].upper() == "DEST" and word[1] > airport_coords[3]:
            dest_coords = (word[0], word[1], word[2], word[3])
            print(f"Координаты 'DEST': x0={dest_coords[0]:.2f}, y0={dest_coords[1]:.2f}")
            break

    if dest_coords is None:
        print("Слово 'DEST' не найдено ниже 'AIRPORT'.")
        # Используем приблизительные координаты
        words_below_airport = [w for w in words if w[1] > airport_coords[3]]
        if words_below_airport:
            min_y0_below = min([w[1] for w in words_below_airport])
            avg_height = 15  # приблизительная высота строки
            YY_airport.append(min_y0_below - 2)
            YY_airport.append(min_y0_below + avg_height + 2)
        else:
            YY_airport.append(airport_coords[3] + 30)
            YY_airport.append(airport_coords[3] + 50)
    else:
        YY_airport.append(dest_coords[1] - 2)
        YY_airport.append(dest_coords[3] + 2)

    print("Массив YY (аэропорты):", [f"{val:.2f}" for val in YY_airport])

    # Извлечение текста
    blocks = page_with_table.get_text("dict").get("blocks", [])
    extracted_text_dict = {}
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text_content = span["text"].strip()
                    bbox = span["bbox"]
                    center_x = (bbox[0] + bbox[2]) / 2
                    center_y = (bbox[1] + bbox[3]) / 2
                    extracted_text_dict[(center_x, center_y)] = text_content

    def find_text_in_rect(x0, y0, x1, y1):
        found_texts = []
        for (cx, cy), text in extracted_text_dict.items():
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                found_texts.append(text)
        return " ".join(found_texts).strip()

    num_cols_airport = len(XX_airport) - 1
    num_rows_airport = len(YY_airport) - 1
    df_data_airport = []

    for j in range(num_rows_airport):
        row_data = []
        for i in range(num_cols_airport):
            x_start = XX_airport[i]
            x_end = XX_airport[i+1]
            y_start = YY_airport[j]
            y_end = YY_airport[j+1]
            cell_text = find_text_in_rect(x_start, y_start, x_end, y_end)
            row_data.append(cell_text)
        df_data_airport.append(row_data)

    return pd.DataFrame(df_data_airport)

def extract_airport_maps(doc):
    """Извлекает карты аэропортов с последней страницы"""
    last_page = doc[-1]
    
    # Извлечение текста построчно
    text_blocks = last_page.get_text("dict")
    lines = []
    for block in text_blocks["blocks"]:
        if "lines" in block:
            for line in block["lines"]:
                line_text = "".join(span["text"] for span in line["spans"])
                stripped = line_text.strip()
                if stripped:
                    lines.append(stripped)

    if len(lines) < 4:
        print("⚠️ На последней странице меньше 4 непустых строк текста!")
        text_A1 = "DEP LFMQ"
        text_A28 = "DEST LFMV"
    else:
        text_A1 = lines[1] if len(lines) > 1 else "DEP"
        text_A28 = lines[3] if len(lines) > 3 else "DEST"

    # Извлечение и обработка изображений
    image_list = last_page.get_images(full=True)
    img_paths = []

    for idx, img in enumerate(image_list):
        try:
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            pil_img = PILImage.open(io.BytesIO(image_bytes))
            pil_img = pil_img.resize((500, 500), PILImage.LANCZOS)
            
            # Сохраняем в память
            img_buffer = io.BytesIO()
            pil_img.save(img_buffer, format='PNG')
            img_paths.append(img_buffer)
        except Exception as e:
            print(f"Ошибка при обработке изображения {idx}: {e}")

    return text_A1, text_A28, img_paths

def create_generated_sheet(wb, ws1, ws2, ws3, ws4):
    """Создает лист Generated_Sheet на основе данных из других листов"""
    new_sheet_name = "Generated_Sheet"
    if new_sheet_name in wb.sheetnames:
        wb.remove(wb[new_sheet_name])
    ws = wb.create_sheet(title=new_sheet_name)

    # Общий шрифт
    default_font = Font(name='Helvetica Neue', size=11)

    # Ширина столбцов
    col_widths = {'A': 3, 'B': 25, 'C': 8, 'D': 8, 'E': 8, 'F': 8, 'G': 8, 'H': 20}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Стили
    bold_gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    bold_font = Font(name='Helvetica Neue', size=11, bold=True)
    header_font = Font(name='Helvetica Neue', size=11, bold=True)

    # Границы
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Заголовки (строки 1–2)
    headers = {
        'A1': ('№', 'center', 'top', True, 2),
        'B1': ('Waypoint', 'left', 'top', True, 2),
        'C1': ('ALT', 'center', 'center', True, 2),
        'D1': ('HDG', 'left', 'center', False, 1),
        'E1': ('Dist.', 'left', 'center', False, 1),
        'F1': ('EFOB', 'left', 'center', False, 1),
        'G1': ('ETA', 'left', 'center', False, 1),
        'H1': ('Radio', 'left', 'top', True, 2),
        'D2': ('CRS', 'right', 'center', False, 1),
        'E2': ('Time', 'right', 'center', False, 1),
        'F2': ('AFOB', 'right', 'center', False, 1),
        'G2': ('ATA', 'right', 'center', False, 1),
    }

    for cell_ref, (value, h_align, v_align, merge, rows) in headers.items():
        cell = ws[cell_ref]
        cell.value = value
        cell.font = header_font
        cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=False)
        if merge:
            col_letter = cell_ref[0]
            start_row = int(cell_ref[1:])
            end_row = start_row + rows - 1
            col_idx = ord(col_letter.upper()) - ord('A') + 1
            ws.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)

    # Фон и жирный шрифт для A1:H2
    for row in range(1, 3):
        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.fill = bold_gray_fill
            cell.font = bold_font
            cell.border = thin_border

    # Переменные
    y0 = 5
    last_row_main = ws2.max_row
    x = last_row_main - 3  # количество циклов

    # Обработка строк
    for i in range(1, x + 1):
        row_offset = y0 + i * 3

        # A: номер с ведущим нулём
        num_val = i + 1
        a_val = f"{num_val:02d}"
        a_cell = ws.cell(row=row_offset, column=1, value=a_val)
        a_cell.font = default_font
        a_cell.alignment = Alignment(horizontal='center', vertical='top')
        if i == x:
            ws.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset+4, end_column=1)
        else:
            ws.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset+2, end_column=1)

        # B: Waypoint
        b_val = ws2.cell(row=i+3, column=1).value
        b_cell = ws.cell(row=row_offset, column=2, value=b_val)
        b_cell.font = default_font
        b_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        if i == x:
            ws.merge_cells(start_row=row_offset, start_column=2, end_row=row_offset+4, end_column=2)
        else:
            ws.merge_cells(start_row=row_offset, start_column=2, end_row=row_offset+2, end_column=2)

        # C: ALT (верхняя часть)
        c_val = ws2.cell(row=i+3, column=5).value
        if c_val:
            c_cell = ws.cell(row=row_offset - 1, column=3, value=c_val)
            c_cell.font = default_font
            c_cell.alignment = Alignment(horizontal='center', vertical='center')
            if i < x:
                ws.merge_cells(start_row=row_offset - 1, start_column=3, end_row=row_offset, end_column=3)

        # D: HDG/CRS
        d1_val = ws2.cell(row=i+3, column=3).value
        if d1_val:
            d1_cell = ws.cell(row=row_offset - 1, column=4, value=d1_val)
            d1_cell.font = default_font
            d1_cell.alignment = Alignment(horizontal='left', vertical='center')

        d2_val = ws2.cell(row=i+3, column=4).value
        if d2_val:
            d2_cell = ws.cell(row=row_offset, column=4, value=d2_val)
            d2_cell.font = default_font
            d2_cell.alignment = Alignment(horizontal='right', vertical='center')

        # E: Dist/Time
        e1_val = ws2.cell(row=i+3, column=11).value
        if e1_val:
            e1_cell = ws.cell(row=row_offset - 1, column=5, value=e1_val)
            e1_cell.font = default_font
            e1_cell.alignment = Alignment(horizontal='left', vertical='center')

        e2_val = ws2.cell(row=i+3, column=16).value
        if e2_val:
            e2_cell = ws.cell(row=row_offset, column=5, value=e2_val)
            e2_cell.font = default_font
            e2_cell.alignment = Alignment(horizontal='right', vertical='center')

        # F: EFOB/AFOB
        f_val = ws2.cell(row=i+3, column=14).value
        if f_val:
            f_cell = ws.cell(row=row_offset - 1, column=6, value=f_val)
            f_cell.font = default_font
            f_cell.alignment = Alignment(horizontal='left', vertical='center')

        # G: ETA/ATA — оставляем пустым

        # H: Radio — объединяем две строки
        h_start = row_offset - 1
        h_end = row_offset
        ws.merge_cells(start_row=h_start, start_column=8, end_row=h_end, end_column=8)
        h_cell = ws.cell(row=h_start, column=8)
        h_cell.font = default_font
        h_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # Последняя строка блока (если не последний блок)
        if i < x:
            merge_range = f"C{row_offset + 1}:H{row_offset + 1}"
            ws.merge_cells(merge_range)
            merged_cell = ws.cell(row=row_offset + 1, column=3)
            merged_cell.font = default_font
            merged_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # Первый блок (A3:A7 и т.д.)
    # A3 = "01", merge A3:A7
    ws.merge_cells('A3:A7')
    a3 = ws['A3']
    a3.value = "01"
    a3.font = default_font
    a3.alignment = Alignment(horizontal='center', vertical='top')

    # B3 = Main_Route_Grid.A3, merge B3:B7
    b3_val = ws2.cell(row=3, column=1).value
    ws.merge_cells('B3:B7')
    b3 = ws['B3']
    b3.value = b3_val
    b3.font = default_font
    b3.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # C3:H6 — большой блок
    ws.merge_cells('C3:H6')
    c3 = ws['C3']

    # Получаем значения с защитой от None
    h2_val = ws3['H2'].value if ws3['H2'].value else "____"
    h2_plus_300 = (h2_val + 300) if isinstance(h2_val, (int, float)) else '____'
    
    # Формирование текста для C3
    text_c3 = (
        f"Departure ({ws4['A1'].value if ws4['A1'].value else '____'}, ____, {h2_val}, _______, "
        f"{h2_plus_300}, ___________;\n"
        f"ATIS: {ws3['D2'].value if ws3['D2'].value else '____'}; GND: {ws3['G2'].value if ws3['G2'].value else '____'}; TWR: {ws3['E2'].value if ws3['E2'].value else '____'};\n"
        f"RWY: {ws3['I2'].value if ws3['I2'].value else '____'}; Length: {ws3['J2'].value if ws3['J2'].value else '____'} m.; Req. Dist.: _____ m.; Surface: _____;\n"
        f"Exp. Wind: {ws1['C5'].value if ws1['C5'].value else '____'}; Exp. QNH: ___ hpa;\n"
        f"RWY: ____ ; Wind: ________; QNH: _______; Squak: ________"
    )

    c3.value = text_c3
    c3.font = Font(name='Helvetica Neue', size=9)
    c3.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # Последний блок
    final_start_row = y0 + x * 3 + 1
    final_end_row = final_start_row + 3
    ws.merge_cells(start_row=final_start_row, start_column=3, end_row=final_end_row, end_column=8)
    final_cell = ws.cell(row=final_start_row, column=3)

    # Формирование финального текста
    h3_val = ws3['H3'].value if ws3['H3'].value else "____"
    h3_plus_300 = (h3_val + 300) if isinstance(h3_val, (int, float)) else '____'

    text_final = (
        f"Departure ({ws4['A28'].value if ws4['A28'].value else '____'}, ____, {h3_val}, _______, "
        f"{h3_plus_300}, ___________;\n"
        f"ATIS: {ws3['D3'].value if ws3['D3'].value else '____'}; GND: {ws3['G3'].value if ws3['G3'].value else '____'}; TWR: {ws3['E3'].value if ws3['E3'].value else '____'}; A/A: ______; Approach: _______;\n"
        f"RWY: {ws3['I3'].value if ws3['I3'].value else '____'}; Length: {ws3['J3'].value if ws3['J3'].value else '____'}; Req. Dist.: _____ m.; Surface: _____;\n"
        f"Exp. Wind: {ws1['C5'].value if ws1['C5'].value else '____'}; Exp. QNH: ___ hpa;\n"
        f"RWY: ____ ; Wind: ________; QNH: _______; Squak: ________"
    )

    final_cell.value = text_final
    final_cell.font = Font(name='Helvetica Neue', size=9)
    final_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # Настройка высоты строк
    last_output_row = y0 + x * 3 + 4
    for row_num in range(1, last_output_row + 1):
        if (3 <= row_num <= 6) or (final_start_row <= row_num <= final_end_row):
            ws.row_dimensions[row_num].height = 15
        else:
            ws.row_dimensions[row_num].height = 14

    # Применяем границы и шрифт
    for row in ws.iter_rows(min_row=1, max_row=last_output_row, min_col=1, max_col=8):
        for cell in row:
            if cell.font.size != 9:
                cell.font = default_font
            cell.border = thin_border

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
    print(f"Takeoff файл: {takeoff_file_name} (используется только для определения)")
    
    # Открываем PDF для обработки
    doc = fitz.open(stream=processing_file_bytes, filetype="pdf")
    
    try:
        # === ЛИСТ 1: ОСНОВНОЕ ===
        print("Создаю лист 'Основное'...")
        lines = extract_first_n_lines_from_doc(doc, n=32)
        while len(lines) < 32:
            lines.append("")

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Основное"

        ws1.cell(row=1, column=1, value=lines[0] if len(lines) > 0 else "")
        ws1.cell(row=2, column=1, value=lines[1] if len(lines) > 1 else "")
        ws1.cell(row=1, column=7, value=lines[2] if len(lines) > 2 else "")
        ws1.cell(row=2, column=7, value=lines[3] if len(lines) > 3 else "")

        block1 = lines[4:18]
        if len(block1) == 14:
            for col in range(7):
                ws1.cell(row=4, column=1 + col, value=block1[col * 2] if col * 2 < len(block1) else "")
                ws1.cell(row=5, column=1 + col, value=block1[col * 2 + 1] if col * 2 + 1 < len(block1) else "")

        block2 = lines[18:32]
        if len(block2) == 14:
            for col in range(7):
                ws1.cell(row=7, column=1 + col, value=block2[col * 2] if col * 2 < len(block2) else "")
                ws1.cell(row=8, column=1 + col, value=block2[col * 2 + 1] if col * 2 + 1 < len(block2) else "")

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

        # Создание листа Main_Route_Grid
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

        # Стилизация и объединение строки 1
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

        # Автоширина
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

        # === ЛИСТ 3: AIRPORT TABLE ===
        print("Парсинг таблицы аэропортов...")
        df_airport = parse_airport_table(doc)
        
        if df_airport.empty:
            print("⚠️ Таблица аэропортов не найдена, создаю пустой лист")
            ws3 = wb.create_sheet(title="Airport_Table")
            ws3['A1'] = "Таблица аэропортов не найдена в документе"
        else:
            # Создание листа Airport_Table
            ws3 = wb.create_sheet(title="Airport_Table")

            headers = ["", "AIRPORT", "ETA", "WX", "TWR/CTAF", "CLR", "GND", "ELEV", "RWY", "LONGEST"]
            if len(headers) > len(df_airport.columns):
                headers = headers[:len(df_airport.columns)]
            elif len(headers) < len(df_airport.columns):
                headers += [""] * (len(df_airport.columns) - len(headers))

            # Заголовки
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            bold_font_yellow = Font(bold=True)

            for col_num, value in enumerate(headers, 1):
                cell = ws3.cell(row=1, column=col_num, value=value)
                cell.font = bold_font_yellow
                cell.fill = yellow_fill

            # Данные
            for r_idx, row in enumerate(dataframe_to_rows(df_airport, index=False, header=False), start=2):
                for c_idx, value in enumerate(row, start=1):
                    ws3.cell(row=r_idx, column=c_idx, value=value)

            # Автоширина для листа 3
            max_col = ws3.max_column
            for col_idx in range(1, max_col + 1):
                max_len = 0
                for row in ws3.iter_rows(min_col=col_idx, max_col=col_idx, min_row=1, max_row=ws3.max_row):
                    cell = row[0]
                    if hasattr(cell, 'value') and cell.value is not None:
                        try:
                            max_len = max(max_len, len(str(cell.value)))
                        except:
                            pass
                adjusted_width = min(max_len + 2, 50)
                ws3.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

        # === ЛИСТ 4: AIRPORT MAPS ===
        print("Извлечение карт аэропортов с последней страницы...")
        text_A1, text_A28, img_paths = extract_airport_maps(doc)
        
        ws4 = wb.create_sheet(title="Airport_Maps")
        ws4.page_margins = PageMargins(left=0.25, right=0.25, top=0.25, bottom=0.25, header=0.1, footer=0.1)

        ws4['A1'] = text_A1
        ws4['A1'].font = Font(bold=True)

        # Вставка изображений
        if len(img_paths) >= 1:
            img_buffer = img_paths[0]
            img_buffer.seek(0)
            ws4.add_image(XLImage(img_buffer), 'A2')
        
        if len(img_paths) >= 2:
            img_buffer = img_paths[1]
            img_buffer.seek(0)
            ws4.add_image(XLImage(img_buffer), 'A29')

        ws4['A28'] = text_A28
        ws4['A28'].font = Font(bold=True)

        ws4.column_dimensions['A'].width = 70

        # === ЛИСТ 5: Generated_Sheet ===
        print("Создание нового листа 'Generated_Sheet'...")
        create_generated_sheet(wb, ws1, ws2, ws3, ws4)

        print(f"✅ Обработка завершена! Создан Excel файл с 5 листами.")
        
    finally:
        # Всегда закрываем документ
        doc.close()

    # Сохраняем workbook в bytes
    excel_bytes = io.BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)
    
    return excel_bytes.getvalue()

# Функция для обратной совместимости
def process(input_path, output_path):
    """Функция для обратной совместимости (не используется для двух файлов)"""
    raise NotImplementedError("Для обработки двух файлов используйте process_two_pdfs")

def main():
    """Для локального тестирования"""
    import sys
    if len(sys.argv) != 3:
        print("Использование: python your_script.py <файл1.pdf> <файл2.pdf>")
        sys.exit(1)
    
    with open(sys.argv[1], 'rb') as f1, open(sys.argv[2], 'rb') as f2:
        result = process_two_pdfs(f1.read(), f2.read(), sys.argv[1], sys.argv[2])
        
    with open("Flight_Log_Extracted.xlsx", 'wb') as f:
        f.write(result)
    print("Файл сохранен как Flight_Log_Extracted.xlsx")

if __name__ == "__main__":
    main()

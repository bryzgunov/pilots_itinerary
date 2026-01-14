import fitz  # PyMuPDF
import pandas as pd
import openpyxl
import tempfile
import os
import io
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill

def process(input_path, output_path):
    """
    Обработка PDF файла - основная функция
    Вход: input_path - путь к входному PDF файлу
          output_path - путь для сохранения Excel файла
    """
    
    print(f"Начинаю обработку: {input_path} -> {output_path}")
    
    # Открываем PDF
    try:
        doc = fitz.open(input_path)
    except Exception as e:
        raise Exception(f"Не удалось открыть PDF файл: {str(e)}")
    
    try:
        page = doc[0]  # Берем первую страницу
        
        # Извлекаем все слова на странице
        all_words = page.get_text("words") # [(x0, y0, x1, y1, text, block_no, line_no, wno_in_line)]
        
        # Найдем приблизительную Y-координату строки заголовка
        target_y = None
        for word_tuple in all_words:
            x0, y0, x1, y1, text, *_ = word_tuple
            # Поиск первого слова "WAYPOINT" в строке, которая также содержит "ACT" близко по горизонтали
            if text == "WAYPOINT":
                # Проверим, находятся ли другие ключевые слова рядом по Y
                act_found_nearby = False
                for w in all_words:
                    wx0, wy0, wx1, wy1, wtext, *_ = w
                    if wtext == "ACT" and abs((y0 + y1)/2 - (wy0 + wy1)/2) < 5: # Допуск 5 пикселей по вертикали
                        if wx0 > x0: # ACT правее WAYPOINT
                            target_y = (y0 + y1) / 2 # Средняя Y координата строки
                            act_found_nearby = True
                            break
                if act_found_nearby:
                    break

        if target_y is None:
            # Попробуем найти по другому, например, ища "MAG" и смещаясь вниз
            for word_tuple in all_words:
                x0, y0, x1, y1, text, *_ = word_tuple
                if text == "MAG":
                    target_y = (y0 + y1) / 2 + 15 # Примерно на 15 пикселей ниже
                    break

        if target_y is None:
            raise ValueError("Не найдена строка заголовка в PDF.")

        # Фильтруем слова, которые находятся на уровне target_y (с допуском)
        header_words_info = []
        tolerance = 5.0  # Допуск по Y
        for word_tuple in all_words:
            x0, y0, x1, y1, text, *_ = word_tuple
            center_y = (y0 + y1) / 2
            if abs(center_y - target_y) <= tolerance and text in ["WAYPOINT", "AIRWAY", "HDG", "CRS", "ALT", "CMP", "DIR/SPD", "ISA", "TAS", "GS", "LEG", "REM", "USED", "ACT", "ETE"]:
                header_words_info.append((text, x0, x1))

        # Сортируем по X координате
        header_words_info.sort(key=lambda item: item[1])

        # Составляем массив XX
        XX = []
        for i in range(1, len(header_words_info)):
            x1_prev = header_words_info[i-1][2] # x1 предыдущего слова
            x0_next = header_words_info[i][1]   # x0 последующего слова
            boundary_x = (x0_next - x1_prev) / 2 + x1_prev
            XX.append(boundary_x)

        # Редактируем массив XX
        if XX: # Проверяем, что массив не пуст
            # Находим x0 AIRWAY
            x0_airway = None
            for text, x0, x1 in header_words_info:
                if text == "AIRWAY":
                    x0_airway = x0
                    break

            if x0_airway is not None:
                # Первый элемент XX делаем равным x0(AIRWAY) - 2
                XX[0] = x0_airway - 2

            # Добавляем слева элемент 5
            XX.insert(0, 5)

            # Добавляем справа элемент, больший последнего на 10
            if XX: # Проверяем снова, на случай если первый элемент был None и XX не изменилось
                last_val = XX[-1]
                new_last_val = last_val + 10
                XX.append(new_last_val)

        # --- Новая логика для YY ---

        # 1. Найти координаты слова "ALT" из header_words_info
        alt_coords = None
        for text, x0, x1 in header_words_info:
            if text == "ALT":
                # Найдем соответствующие y0, y1 для этого x0, x1 среди all_words
                for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
                    if wtext == "ALT" and abs(wx0 - x0) < 1 and abs(wx1 - x1) < 1: # Строгое совпадение X координат
                        alt_coords = (wx0, wy0, wx1, wy1)
                        break
                if alt_coords:
                     break

        if not alt_coords:
            raise ValueError("Не найдены точные координаты для слова 'ALT'.")

        x0_alt, y0_alt, x1_alt, y1_alt = alt_coords

        # 2. Найти слово "ALTERNATE" и его координату y0
        y0_alternate = None
        for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
            if "ALTERNATE" in wtext:
                y0_alternate = wy0
                break

        if y0_alternate is None:
            raise ValueError("Не найдено слово 'ALTERNATE'.")

        # 3. Найти слова между y1_alt и y0_alternate в области x0_alt - x1_alt
        YY = [] # Массив для y2

        for wx0, wy0, wx1, wy1, wtext, *_ in all_words:
            # Проверяем, находится ли слово внутри области по X и между строками по Y
            if x0_alt <= (wx0 + wx1) / 2 <= x1_alt and y1_alt <= wy0 <= y0_alternate:
                # Проверяем, что слово не является "ALT" или "ALTERNATE" (или их частью)
                if wtext != "ALT" and "ALTERNATE" not in wtext:
                    y2 = wy0 - 2 # Вычисляем y2
                    YY.append(y2) # Сохраняем в массив

        # 4. Добавить y2 слова "ALTERNATE" в конец массива
        y2_alternate = y0_alternate - 2
        YY.append(y2_alternate)

        # --- Парсинг данных по сетке XX, YY (Новая логика) ---

        num_cols = len(XX) - 1  # Количество столбцов = количество интервалов между границами X
        num_rows = len(YY) - 1  # Количество строк = количество интервалов между границами Y

        # Инициализируем DataFrame
        # Определим имена столбцов жестко, так как XX не отражает их напрямую после добавления 5 и +10
        column_names = ['WAYPOINT', 'AIRWAY', 'HDG', 'CRS', 'ALT', 'CMP', 'DIR/SPD', 'ISA', 'TAS', 'GS', 'LEG', 'REM', 'USED', 'REM', 'ACT', 'LEG', 'REM', 'ETE', 'ACT']
        # Обрежем или дополним имена столбцов до num_cols
        if len(column_names) < num_cols:
            for i in range(len(column_names), num_cols):
                column_names.append(f'COL_{i}')
        elif len(column_names) > num_cols:
            column_names = column_names[:num_cols]

        data_grid = [['' for _ in range(num_cols)] for _ in range(num_rows)]

        # Проходим по сетке
        for row_idx in range(num_rows):
            for col_idx in range(num_cols):
                # Определяем границы ячейки
                x_min = XX[col_idx]
                x_max = XX[col_idx + 1]
                y_min = YY[row_idx]
                y_max = YY[row_idx + 1]

                # Ищем слова, центры которых попадают в эту ячейку
                cell_texts = []
                for word_tuple in all_words:
                    wx0, wy0, wx1, wy1, wtext, *_ = word_tuple
                    center_x = (wx0 + wx1) / 2
                    center_y = (wy0 + wy1) / 2

                    if x_min <= center_x <= x_max and y_min <= center_y <= y_max:
                        cell_texts.append(wtext)

                # Объединяем найденные слова в одну строку
                if cell_texts:
                    # Пробел между словами может быть важен, например, для "0:04 0:31 0:04"
                    # или "H3 332/005". Используем пробел как разделитель.
                    combined_text = ' '.join(cell_texts)
                    data_grid[row_idx][col_idx] = combined_text

        df = pd.DataFrame(data_grid, columns=column_names)
        
        print(f"✅ Парсинг сетки завершен! DataFrame создан: {df.shape[0]} строк, {df.shape[1]} столбцов.")

        # --- Сохранение в Excel ---
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Main_Route_Grid_Parsed', index=False)
            worksheet = writer.sheets['Main_Route_Grid_Parsed']

            # Определяем стили
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            align_center = Alignment(horizontal="center", vertical="center")

            # Форматируем заголовки
            for col_num, value in enumerate(df.columns.values, start=1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center

            # Авто-ширина колонок
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"💾 Таблица сохранена как: {output_path}")
        
        # Закрываем документ
        doc.close()
        
        return True
        
    except Exception as e:
        # Закрываем документ в случае ошибки
        doc.close()
        raise e

def main():
    """
    Точка входа для тестирования скрипта локально
    """
    import sys
    
    if len(sys.argv) != 3:
        print("Использование: python your_script.py <входной_файл.pdf> <выходной_файл.xlsx>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if not os.path.exists(input_file):
        print(f"Ошибка: файл не найден: {input_file}")
        sys.exit(1)
    
    try:
        success = process(input_file, output_file)
        if success:
            print("✅ Обработка завершена успешно!")
            sys.exit(0)
        else:
            print("❌ Ошибка при обработке файла")
            sys.exit(1)
    except Exception as e:
        print(f"❌ Ошибка: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()

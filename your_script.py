# your_script.py - демо-скрипт для тестирования
import os
import time

def process(input_path, output_path):
    """
    Простая функция обработки файла для демонстрации
    """
    print(f"Начинаю обработку: {input_path} -> {output_path}")
    
    # Проверяем, существует ли файл
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Файл не найден: {input_path}")
    
    # Определяем тип файла по расширению
    file_ext = os.path.splitext(input_path)[1].lower()
    
    # Для текстовых файлов
    if file_ext in ['.txt', '.csv', '.json', '.xml', '.html', '.py', '.js']:
        try:
            # Пытаемся прочитать как текст
            with open(input_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Простая обработка текста
            processed = f"""=== НАЧАЛО ОБРАБОТАННОГО ФАЙЛА ===

Исходный файл: {os.path.basename(input_path)}
Время обработки: {time.strftime('%Y-%m-%d %H:%M:%S')}
Размер исходного файла: {os.path.getsize(input_path)} байт

{content}

=== КОНЕЦ ОБРАБОТАННОГО ФАЙЛА ===
Обработано скриптом your_script.py
"""
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(processed)
                
            print(f"Текстовый файл обработан: {output_path}")
            
        except UnicodeDecodeError:
            # Если не удалось прочитать как текст, обрабатываем как бинарный
            process_binary(input_path, output_path)
    
    # Для бинарных файлов
    else:
        process_binary(input_path, output_path)

def process_binary(input_path, output_path):
    """Обработка бинарных файлов"""
    with open(input_path, 'rb') as f_in, open(output_path, 'wb') as f_out:
        # Читаем файл
        data = f_in.read()
        
        # Добавляем заголовок
        header = f"=== БИНАРНЫЙ ФАЙЛ ===\n"
        header += f"Имя: {os.path.basename(input_path)}\n"
        header += f"Размер: {len(data)} байт\n"
        header += f"Время: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        header += "===================\n\n"
        
        # Пишем заголовок и данные
        f_out.write(header.encode('utf-8'))
        f_out.write(data)
        
        # Добавляем футер
        footer = b"\n\n=== КОНЕЦ ФАЙЛА ==="
        f_out.write(footer)
    
    print(f"Бинарный файл обработан: {output_path}")

# Для запуска из командной строки
if __name__ == "__main__":
    import sys
    if len(sys.argv) == 3:
        process(sys.argv[1], sys.argv[2])
        print("✅ Обработка завершена!")
    else:
        print("Использование: python your_script.py <входной_файл> <выходной_файл>")

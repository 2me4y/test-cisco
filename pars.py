import json
import re
from docx import Document

# Список всех возможных красных цветов в Word (HEX коды)
# FF0000 - Стандартный красный
# C00000 - Темно-красный (часто используется в тестах)
# FF4B4B - Ваш вариант
# E03E2D - Еще один оттенок
RED_COLORS = ["FF0000", "C00000", "FF4B4B", "E03E2D", "FF3333"]

def is_run_colored_red(paragraph):
    """
    Проверяет фрагменты текста на наличие красного цвета из списка.
    """
    for run in paragraph.runs:
        font = run.font
        
        # 1. Проверка цвета шрифта (Font Color)
        if font.color and font.color.rgb:
            color_hex = str(font.color.rgb).upper()
            if color_hex in RED_COLORS:
                print(f"   [!] Найден ответ (Цвет {color_hex}): {run.text.strip()}")
                return True
                
        # 2. (Опционально) Проверка выделения маркером (Highlight)
        # Иногда "цвет" это не шрифт, а фон. В Word красный маркер часто имеет индекс 6 или 7.
        # Если нужно, раскомментируйте:
        # if font.highlight_color:
        #     return True
            
    return False

def parse_docx_smart_color(docx_path, json_path):
    document = Document(docx_path)
    questions = []
    current_question = None
    
    # Регулярки (как в прошлом скрипте)
    question_pattern = re.compile(r'^\d+\s*[-.)]\s+')
    traps_pattern = re.compile(r'^\d+-to-[A-Za-z0-9]', re.IGNORECASE)

    print(f"Начинаю парсинг файла: {docx_path}")

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text: continue

        is_new_question = False
        if traps_pattern.match(text):
            is_new_question = False
        elif question_pattern.match(text):
            is_new_question = True

        if is_new_question:
            if current_question:
                questions.append(current_question)
            
            current_question = {
                "question": text,
                "options": [],
                "correct_index": -1,
                "correct_answer": ""
            }
            
        elif current_question:
            # Это вариант ответа
            current_question["options"].append(text)
            
            # Проверка цвета
            if is_run_colored_red(paragraph):
                idx = len(current_question["options"]) - 1
                current_question["correct_index"] = idx
                current_question["correct_answer"] = text

    if current_question:
        questions.append(current_question)

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, indent=4)
        
    print(f"Готово! Сохранено вопросов: {len(questions)}")

# --- ЗАПУСК ---
input_file = "test_questions.docx" # Проверьте имя файла!
output_file = "output_red_fixed.json"

try:
    parse_docx_smart_color(input_file, output_file)
except Exception as e:
    print(f"Ошибка: {e}")
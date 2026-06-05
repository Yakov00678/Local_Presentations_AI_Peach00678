import os
import re
from pptx import Presentation
from pptx.util import Inches
from designer import apply_slide_design

def parse_markdown_to_slides(text):
    raw_slides = re.split(r'#\s*СЛАЙД\s*\d+', text)
    slides_data = []
    
    for raw in raw_slides:
        lines = [line.strip() for line in raw.split('\n') if line.strip()]
        if not lines:
            continue
            
        slide_info = {'title': '', 'content_lines': [], 'table': None}
        table_rows = []
        
        for line in lines:
            if line.startswith('---'):
                continue
            if line.startswith('##'):
                slide_info['title'] = line.replace('##', '').strip()
            elif line.startswith('|'):
                if set(line.replace('|', '').replace('-', '').replace(' ', '')) <= {' '}:
                    continue
                cells = [c.strip() for c in line.split('|')][1:-1]
                if cells:
                    table_rows.append(cells)
            else:
                slide_info['content_lines'].append(line)
                
        if table_rows:
            slide_info['table'] = table_rows
            
        if slide_info['title'] or slide_info['content_lines'] or slide_info['table']:
            slides_data.append(slide_info)
        
    return slides_data

def get_unique_filename(folder, base_name, extension):
    """Функция подбирает уникальное имя файла, добавляя цифру в конец, если файл уже существует"""
    # Сначала проверяем базовое имя: presentation.pptx
    filename = f"{base_name}{extension}"
    file_path = os.path.join(folder, filename)
    
    counter = 1
    # Если файл существует, крутим цикл, пока не найдем свободное имя
    while os.path.exists(file_path):
        filename = f"{base_name}_{counter}{extension}"
        file_path = os.path.join(folder, filename)
        counter += 1
        
    return file_path

def create_presentation(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    for data in slides_data:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # Вызов темы оформления (можно поменять на "dark_style")
        apply_slide_design(slide, data, theme_name="light_premium")

    # НАСТРОЙКА СОХРАНЕНИЯ С АВТОНУМЕРАЦИЕЙ
    output_folder = "Presentations"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    # Задаем базовое имя файла (без цифры) и расширение
    base_name = "presentation"
    extension = ".pptx"
    
    # Получаем уникальный путь (например, Presentations/presentation_1.pptx)
    file_path = get_unique_filename(output_folder, base_name, extension)
    
    prs.save(file_path)
    print(f"\n✅ Презентация успешно сохранена по пути: {file_path}")

# Ваша входная разметка
markdown_input = """
# СЛАЙД 2
## Интервью: анатомия диалога
Интервью — это не «просто поговорить». Это инструмент:
- Получения информации
- Формирования общественного мнения

---
# СЛАЙД 3
## Типология интервью


| Тип | Цель | Характеристика |
|-----|------|----------------|
| Блиц | Быстрая фиксация | Короткие вопросы |
| Блиц | Быстрая фиксация | Короткие вопросы |
| Блиц | Быстрая фиксация | Короткие вопросы |
"""

slides = parse_markdown_to_slides(markdown_input)
create_presentation(slides)

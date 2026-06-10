from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE

# НАШ СЛОВАРЬ ЦВЕТОВЫХ ТЕМ
THEMES = {
    "1": { # 1 - Dark (ПОЛНОСТЬЮ ТЕМНЫЙ СТИЛЬ)
        "bg": RGBColor(24, 28, 36),                  # Глубокий темный фон слайда
        "text": RGBColor(240, 244, 248),             # Мягкий белый текст
        "accent": RGBColor(52, 152, 219),            # Яркий неоново-синий для акцентных линий
        "table_header_bg": RGBColor(34, 40, 52),     # Темная стальная шапка
        "table_header_text": RGBColor(52, 152, 219),   # Неоново-синий текст в шапке
        "table_row_even": RGBColor(28, 34, 44),      # Четные строки таблицы (чуть светлее)
        "table_row_odd": RGBColor(22, 26, 34),       # Нечетные строки таблицы (чуть темнее)
        "font_title": "Dela Gothic One",             # Кастомный шрифт для заголовков
        "font_body": "Montserrat"                    # Читаемый шрифт для основного контента
    },
    "2": { # 2 - White
        "bg": RGBColor(255, 255, 255),
        "text": RGBColor(10, 10, 10),
        "accent": RGBColor(0, 0, 0),
        "table_header_bg": RGBColor(240, 240, 240),
        "table_header_text": RGBColor(0, 0, 0),
        "table_row_even": RGBColor(255, 255, 255),
        "table_row_odd": RGBColor(248, 248, 248),
        "font_title": "Arial",
        "font_body": "Arial"
    },
    "3": { # 3 - Loft
        "bg": RGBColor(240, 234, 224),
        "text": RGBColor(60, 50, 45),
        "accent": RGBColor(160, 82, 45),
        "table_header_bg": RGBColor(110, 90, 80),
        "table_header_text": RGBColor(240, 234, 224),
        "table_row_even": RGBColor(245, 240, 232),
        "table_row_odd": RGBColor(235, 228, 216),
        "font_title": "Georgia",
        "font_body": "Palatino"
    },
    "4": { # 4 - Neon/Cyberpunk
        "bg": RGBColor(10, 5, 20),
        "text": RGBColor(0, 255, 240),
        "accent": RGBColor(255, 0, 128),
        "table_header_bg": RGBColor(255, 0, 128),
        "table_header_text": RGBColor(10, 5, 20),
        "table_row_even": RGBColor(20, 10, 35),
        "table_row_odd": RGBColor(12, 6, 24),
        "font_title": "Impact",
        "font_body": "Courier New"
    },
    "5": { # 5 - Classic
        "bg": RGBColor(245, 247, 250),
        "text": RGBColor(15, 32, 67),
        "accent": RGBColor(212, 175, 55),
        "table_header_bg": RGBColor(15, 32, 67),
        "table_header_text": RGBColor(255, 255, 255),
        "table_row_even": RGBColor(255, 255, 255),
        "table_row_odd": RGBColor(235, 240, 245),
        "font_title": "Times New Roman",
        "font_body": "Times New Roman"
    },
    "6": { # 6 - Sketch
        "bg": RGBColor(250, 250, 245),
        "text": RGBColor(40, 50, 70),
        "accent": RGBColor(190, 190, 190),
        "table_header_bg": RGBColor(220, 225, 235),
        "table_header_text": RGBColor(40, 50, 70),
        "table_row_even": RGBColor(250, 250, 245),
        "table_row_odd": RGBColor(240, 243, 248),
        "font_title": "Comic Sans MS",
        "font_body": "Comic Sans MS"
    }
}

def apply_slide_design(slide, data, theme_name="1"):
    theme = THEMES.get(theme_name, THEMES["1"])
    
    # 1. Заливка фона слайда
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = theme["bg"]
    
    # Стартовая координата для контента (изменится, если будет заголовок)
    current_top = Inches(2.2)
    
    # 2. Оформление заголовка слайда (с умным отслеживанием длинного текста)
    if data.get('title'):
        title_len = len(data['title'])
        is_long_title = title_len > 30
        
        # Меняем высоту бокса в зависимости от количества символов
        title_box_height = Inches(1.6) if is_long_title else Inches(0.9)
        
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), title_box_height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data['title']
        p.font.name = theme["font_title"]
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = theme["text"]
        
        # Динамическая позиция линии: сдвигаем ниже, если заголовок перенесся на 2 строки
        line_top = Inches(2.1) if is_long_title else Inches(1.5)
        
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), line_top, Inches(3.0), Inches(0.04)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = theme["accent"]
        line.line.color.rgb = theme["accent"]
        
        # Обновляем координату начала контента под линией заголовка
        current_top = line_top + Inches(0.5)

    # Максимальная высота контентной зоны на слайде
    max_available_height = Inches(4.5)
        
    # 3. Оформление текста и списков (с рабочим Autofit переполнения)
    if data.get('content_lines'):
        # Если есть таблица — выделяем под текст верхнюю половину экрана, если нет — всё доступное место
        box_height = max_available_height if not data.get('table') else Inches(1.8)
        
        txBox = slide.shapes.add_textbox(Inches(0.8), current_top, Inches(11.5), box_height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE 
        
        for i, line in enumerate(data['content_lines']):
            if not line.strip():
                continue
            
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            
            if line.strip().startswith('-') or line.strip().startswith('•'):
                clean_line = line.replace('-', '').replace('•', '').strip()
                p.text = f"•  {clean_line}"
                p.font.size = Pt(20)
                p.space_before = Pt(6)
            else:
                p.text = line
                p.font.size = Pt(22)
                p.space_before = Pt(12)
                
            p.font.name = theme["font_body"]
            p.font.color.rgb = theme["text"]
            
        # Сдвигаем маркер высоты для таблицы строго под текстовый блок
        current_top += box_height + Inches(0.3)
                
    # 4. Оформление таблицы (с автоматическим подбором ширины колонок)
    if data.get('table'):
        rows = len(data['table'])
        cols = max(len(row) for row in data['table']) if data['table'] else 0
        
        if rows > 0 and cols > 0:
            left = Inches(0.8)
            top = current_top if data.get('content_lines') else Inches(2.2)
            total_width = Inches(11.7)  # Общая ширина таблицы на слайде
            height = Inches(0.4 * rows)
            
            table_shape = slide.shapes.add_table(rows, cols, left, top, total_width, height)
            table = table_shape.table
            
            # --- УМНЫЙ ПОДБОР ШИРИНЫ КОЛОНОК ---
            max_chars_per_col = [0] * cols
            for row in data['table']:
                for c_idx, val in enumerate(row):
                    if c_idx < cols:
                        max_chars_per_col[c_idx] = max(max_chars_per_col[c_idx], len(str(val)))
            
            max_chars_per_col = [max(1, count) for count in max_chars_per_col]
            total_chars = sum(max_chars_per_col)
            
            for c_idx in range(cols):
                col_share = max_chars_per_col[c_idx] / total_chars
                table.columns[c_idx].width = int(total_width * col_share)
            # ----------------------------------
            
            for r_idx, row in enumerate(data['table']):
                for c_idx, val in enumerate(row):
                    if c_idx < len(table.columns):
                        cell = table.cell(r_idx, c_idx)
                        cell.text = val
                        cell.fill.solid()
                        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                        
                        if r_idx == 0:  # Шапка
                            cell.fill.fore_color.rgb = theme["table_header_bg"]
                            text_color = theme["table_header_text"]
                            is_bold = True
                            font_size = Pt(16)
                        elif r_idx % 2 == 0:  # Четная строка
                            cell.fill.fore_color.rgb = theme["table_row_even"]
                            text_color = theme["text"]
                            is_bold = False
                            font_size = Pt(14)
                        else:  # Нечетная строка
                            cell.fill.fore_color.rgb = theme["table_row_odd"]
                            text_color = theme["text"]
                            is_bold = False
                            font_size = Pt(14)
                        
                        for p in cell.text_frame.paragraphs:
                            p.alignment = PP_ALIGN.LEFT
                            p.font.name = theme["font_body"]
                            p.font.size = font_size
                            p.font.color.rgb = text_color
                            p.font.bold = is_bold

import os
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE

# НАШ ОБНОВЛЕННЫЙ СЛОВАРЬ ЦВЕТОВЫХ ТЕМ С ДИЗАЙНЕРСКИМИ ШРИФТАМИ
THEMES = {
    "1": { # 1 - Dark (УЛЬТРАСОВРЕМЕННЫЙ СТИЛЬ)
        "bg": RGBColor(24, 28, 36),
        "text": RGBColor(240, 244, 248),
        "accent": RGBColor(52, 152, 219),
        "table_header_bg": RGBColor(34, 40, 52),
        "table_header_text": RGBColor(52, 152, 219),
        "table_row_even": RGBColor(28, 34, 44),
        "table_row_odd": RGBColor(22, 26, 34),
        "font_title": "Dela Gothic One",             
        "font_body": "Montserrat"                    
    },
    "2": { # 2 - White (ШВЕЙЦАРСКИЙ МИНИМАЛИЗМ)
        "bg": RGBColor(255, 255, 255),
        "text": RGBColor(15, 17, 23),
        "accent": RGBColor(0, 0, 0),
        "table_header_bg": RGBColor(242, 244, 247),
        "table_header_text": RGBColor(0, 0, 0),
        "table_row_even": RGBColor(255, 255, 255),
        "table_row_odd": RGBColor(248, 250, 252),
        "font_title": "Inter",                       
        "font_body": "Arial"
    },
    "3": { # 3 - Loft (ИЗДАТЕЛЬСКИЙ КРАФТ)
        "bg": RGBColor(242, 236, 226),
        "text": RGBColor(61, 51, 46),
        "accent": RGBColor(166, 85, 46),
        "table_header_bg": RGBColor(112, 92, 82),
        "table_header_text": RGBColor(242, 236, 226),
        "table_row_even": RGBColor(247, 242, 234),
        "table_row_odd": RGBColor(237, 230, 218),
        "font_title": "Georgia",                     
        "font_body": "Palatino"
    },
    "4": { # 4 - Neon/Cyberpunk (ЦИФРОВОЙ КОД)
        "bg": RGBColor(11, 6, 21),
        "text": RGBColor(0, 255, 240),
        "accent": RGBColor(255, 0, 128),
        "table_header_bg": RGBColor(255, 0, 128),
        "table_header_text": RGBColor(11, 6, 21),
        "table_row_even": RGBColor(21, 11, 36),
        "table_row_odd": RGBColor(13, 7, 25),
        "font_title": "Impact",                      
        "font_body": "Courier New"                   
    },
    "5": { # 5 - Classic (ПРЕМИАЛЬНЫЙ БИЗНЕС)
        "bg": RGBColor(246, 248, 251),
        "text": RGBColor(16, 33, 68),
        "accent": RGBColor(213, 176, 56),
        "table_header_bg": RGBColor(16, 33, 68),
        "table_header_text": RGBColor(255, 255, 255),
        "table_row_even": RGBColor(255, 255, 255),
        "table_row_odd": RGBColor(236, 241, 246),
        "font_title": "Garamond",                    
        "font_body": "Times New Roman"
    },
    "6": { # 6 - Sketch (БЛОКНОТНЫЙ НАБРОСОК)
        "bg": RGBColor(251, 251, 246),
        "text": RGBColor(41, 51, 71),
        "accent": RGBColor(191, 191, 191),
        "table_header_bg": RGBColor(221, 226, 236),
        "table_header_text": RGBColor(41, 51, 71),
        "table_row_even": RGBColor(251, 251, 246),
        "table_row_odd": RGBColor(241, 244, 249),
        "font_title": "Segoe Print",                 
        "font_body": "Segoe UI"
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
        
        line_top = Inches(2.1) if is_long_title else Inches(1.5)
        
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), line_top, Inches(3.0), Inches(0.04)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = theme["accent"]
        line.line.color.rgb = theme["accent"]
        
        current_top = line_top + Inches(0.5)

    max_available_height = Inches(4.5)
    
    # Детекция картинки: если файл существует, ужимаем текстовую сетку влево
    has_image = data.get('image_path') and os.path.exists(data['image_path'])
    content_width = Inches(6.0) if has_image else Inches(11.5)
        
    # 3. Оформление текста и списков (ТВОЙ НАСТРОЕННЫЙ КОД)
    if data.get('content_lines'):
        box_height = max_available_height if not data.get('table') and not data.get('quote') and not data.get('metrics') else Inches(1.8)
        
        txBox = slide.shapes.add_textbox(Inches(0.8), current_top, content_width, box_height)
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
            
        current_top += box_height + Inches(0.3)

    # 4. ДОБАВЛЕНИЕ КАРТИНКИ (справа, если она есть на диске)
    if has_image:
        img_left = Inches(7.5)
        img_top = Inches(2.2)
        img_width = Inches(5.0)
        slide.shapes.add_picture(data['image_path'], img_left, img_top, width=img_width)

    # 5. КРУПНЫЕ ЦИФРЫ-ПОКАЗАТЕЛИ (МЕТРИКИ)
    if data.get('metrics'):
        metrics_count = len(data['metrics'])
        if metrics_count > 0:
            card_width = content_width / metrics_count - Inches(0.2)
            
            for m_idx, metric in enumerate(data['metrics']):
                m_left = Inches(0.8) + m_idx * (card_width + Inches(0.2))
                m_box = slide.shapes.add_textbox(m_left, current_top, card_width, Inches(1.8))
                tf = m_box.text_frame
                tf.word_wrap = True
                
                # Огромная акцентная цифра
                p_num = tf.paragraphs[0]
                p_num.text = metric['number']
                p_num.font.name = theme["font_title"]
                p_num.font.size = Pt(44)
                p_num.font.bold = True
                p_num.font.color.rgb = theme["accent"]
                
                # Описание под ней
                if metric['description']:
                    p_desc = tf.add_paragraph()
                    p_desc.text = metric['description']
                    p_desc.font.name = theme["font_body"]
                    p_desc.font.size = Pt(14)
                    p_desc.font.color.rgb = theme["text"]
                    p_desc.space_before = Pt(4)
                    
            current_top += Inches(2.0)

    # 6. ОФОРМЛЕНИЕ ЦИТАТЫ (Стильная фоновая плашка-карточка)
    if data.get('quote'):
        quote_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), current_top, content_width, Inches(1.2))
        quote_box.fill.solid()
        quote_box.fill.fore_color.rgb = theme["table_row_even"]
        quote_box.line.color.rgb = theme["accent"]
        
        tf = quote_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = f"“{data['quote']}”"
        p.alignment = PP_ALIGN.LEFT
        p.font.name = theme["font_body"]
        p.font.size = Pt(16)
        p.font.color.rgb = theme["text"]
        p.font.italic = True
        
        current_top += Inches(1.5)
                
     # 7. Оформление таблицы (ТВОЙ НАСТРОЕННЫЙ КОД С АВТОПОДБОРОМ ШИРИНЫ КОЛОНОК)
    if data.get('table'):
        rows = len(data['table'])
        cols = max(len(row) for row in data['table']) if data['table'] else 0
        
        if rows > 0 and cols > 0:
            left = Inches(0.8)
            top = current_top if data.get('content_lines') else Inches(2.2)
            total_width = Inches(11.7)  
            height = Inches(0.4 * rows)
            
            table_shape = slide.shapes.add_table(rows, cols, left, top, total_width, height)
            table = table_shape.table
            
            # Алгоритм автоподбора ширины колонок
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
            
            # Стилизация и заполнение ячеек данными
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

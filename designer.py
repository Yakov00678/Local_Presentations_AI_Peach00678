from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# НАСТРОЙКА ЦВЕТОВЫХ ТЕМ (Стили)
THEMES = {
    "light_premium": {
        "bg": RGBColor(248, 249, 250),        # Светло-серый
        "text": RGBColor(44, 62, 80)         # Графитовый
    },
    "dark_style": {
        "bg": RGBColor(30, 30, 30),          # Почти черный
        "text": RGBColor(240, 240, 240)       # Белый
    }
}

# Акцентные цвета
COLOR_ACCENT = RGBColor(230, 126, 34)   # Оранжевая линия
COLOR_BLUE = RGBColor(41, 128, 185)     # Шапка таблицы

def apply_slide_design(slide, data, theme_name="light_premium"):
    """Функция полностью красит слайд и расставляет элементы по местам"""
    theme = THEMES.get(theme_name, THEMES["light_premium"])
    
    # 1. Заливка фона
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = theme["bg"]
    
    # 2. Оформление заголовка
    if data['title']:
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(1))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data['title']
        p.font.name = 'Trebuchet MS'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = theme["text"]
        
        # Декоративная линия под заголовком
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.0), Inches(0.04)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_ACCENT
        line.line.color.rgb = COLOR_ACCENT
        
    # 3. Оформление текста и списков
    if data['content_lines']:
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.5), Inches(4.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        
        for i, line in enumerate(data['content_lines']):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            
            if line.startswith('-'):
                p.text = "   " + line.replace('-', '').strip()
                p.font.size = Pt(20)
                p.space_before = Pt(10)
                p.level = 1 
            else:
                p.text = line
                p.font.size = Pt(22)
                p.space_before = Pt(16)
                
            p.font.name = 'Calibri'
            p.font.color.rgb = theme["text"]
                
    # 4. Оформление таблицы
    if data['table']:
        rows = len(data['table'])
        cols = max(len(row) for row in data['table']) if data['table'] else 0
        
        if rows > 0 and cols > 0:
            left = Inches(0.8)
            top = Inches(2.4)
            width = Inches(11.7)
            height = Inches(0.5 * rows)
            
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
            table = table_shape.table
            
            for r_idx, row in enumerate(data['table']):
                for c_idx, val in enumerate(row):
                    if c_idx < len(table.columns):
                        cell = table.cell(r_idx, c_idx)
                        cell.text = val
                        
                        for p in cell.text_frame.paragraphs:
                            p.font.name = 'Calibri'
                            p.font.size = Pt(16)
                            p.font.color.rgb = theme["text"]
                            
                            if r_idx == 0:
                                p.font.bold = True
                                p.font.color.rgb = RGBColor(255, 255, 255)
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = COLOR_BLUE

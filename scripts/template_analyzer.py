import docx
from docx.shared import Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING

def analyze_document_template(template_path):
    """Анализ шаблона документа"""
    document = docx.Document(template_path)
    
    # Анализ секций и полей
    sections = document.sections
    section = sections[0]
    
    print("📏 Параметры страницы:")
    print(f"Высота страницы: {section.page_height.mm} мм")
    print(f"Ширина страницы: {section.page_width.mm} мм")
    print(f"Левое поле: {section.left_margin.mm} мм")
    print(f"Правое поле: {section.right_margin.mm} мм")
    print(f"Верхнее поле: {section.top_margin.mm} мм")
    print(f"Нижнее поле: {section.bottom_margin.mm} мм")
    
    # Анализ стилей
    print("\n🔤 Стили документа:")
    for style in document.styles:
        if style.type == 1:  # Paragraph style
            print(f"Стиль: {style.name}")
            if style.paragraph_format:
                print(f"  Выравнивание: {style.paragraph_format.alignment}")
                print(f"  Межстрочный интервал: {style.paragraph_format.line_spacing}")
    
    # Анализ шрифтов
    print("\n🖋️ Шрифты:")
    font_stats = {}
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if run.font.name:
                font_stats[run.font.name] = font_stats.get(run.font.name, 0) + 1
            if run.font.size:
                print(f"Размер шрифта: {run.font.size.pt} пт")
    
    print("\nИспользованные шрифты:")
    for font, count in font_stats.items():
        print(f"{font}: {count} раз")

def main():
    template_path = '/home/user/study/diplom/reference_template.docx'
    analyze_document_template(template_path)

if __name__ == '__main__':
    main()

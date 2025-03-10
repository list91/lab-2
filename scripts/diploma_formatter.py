import os
import re
from typing import List
import markdown
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION

class DiplomaFormatter:
    CHAPTER_TRANSLATIONS = {
        '1_introduction': '1. Введение',
        '2_theoretical_part': '2. Теоретическая часть',
        '3_practical_implementation': '3. Практическая реализация',
        '4_research_methodology': '4. Методология исследования',
        '5_research_results': '5. Результаты исследования',
        '6_practical_significance': '6. Практическая значимость',
        '7_development_prospects': '7. Перспективы развития',
        '8_appendices': '8. Приложения'
    }

    def __init__(self, chapters_dir: str, output_path: str):
        self.chapters_dir = chapters_dir
        self.output_path = output_path
        self.document = Document()
        self._setup_document_styles()

    def _setup_document_styles(self):
        """Настройка стилей документа по ГОСТ"""
        sections = self.document.sections
        for section in sections:
            section.page_height = Mm(297)
            section.page_width = Mm(210)
            section.left_margin = Mm(30)
            section.right_margin = Mm(10)
            section.top_margin = Mm(20)
            section.bottom_margin = Mm(20)

        # Проверка и настройка стиля Normal
        try:
            style = self.document.styles['Normal']
        except KeyError:
            style = self.document.styles.add_style('Normal', WD_STYLE_TYPE.PARAGRAPH)
        
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        style.paragraph_format.first_line_indent = Mm(12.5)

        # Стили для заголовков
        for level in range(1, 4):
            try:
                heading_style = self.document.styles[f'Heading{level}']
            except KeyError:
                heading_style = self.document.styles.add_style(f'Heading{level}', WD_STYLE_TYPE.PARAGRAPH)
            
            heading_style.font.name = 'Times New Roman'
            heading_style.font.size = Pt(14 + (level * 2))
            heading_style.font.bold = True
            heading_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            heading_style.paragraph_format.space_after = Pt(12)

    def _convert_markdown_to_docx(self, markdown_text: str):
        """Конвертация Markdown в docx с сохранением структуры"""
        # Обработка списков
        markdown_text = re.sub(r'^- ', '• ', markdown_text, flags=re.MULTILINE)
        
        # Обработка кода
        markdown_text = re.sub(r'```(.*?)```', r'[Листинг кода]\n\1', markdown_text, flags=re.DOTALL)
        
        html = markdown.markdown(markdown_text)
        
        # Разбор HTML и добавление в документ
        paragraphs = re.split(r'<p>|</p>', html)[1:-1]
        
        for paragraph in paragraphs:
            # Удаление HTML-тегов
            clean_text = re.sub(r'<[^>]+>', '', paragraph).strip()
            
            if clean_text:
                para = self.document.add_paragraph(clean_text, style='Normal')

    def _process_chapter(self, chapter_path: str):
        """Обработка главы"""
        with open(chapter_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Определение имени главы
        path_parts = chapter_path.split('/')
        
        # Проверяем структуру пути
        if path_parts[-1] == 'content.md':
            # Определяем, находится ли файл непосредственно в директории главы или в подкаталоге
            if 'content.md' in path_parts[-1] and path_parts[-2] in ['6_practical_significance', '7_development_prospects', '8_appendices']:
                chapter_name = path_parts[-2]  # Берем имя директории главы
            else:
                # Для подразделов (например, 1.1, 2.1, и т.д.)
                chapter_name = path_parts[-3]  # Берем имя директории главы
        
        # Перевод заголовка главы
        translated_name = self.CHAPTER_TRANSLATIONS.get(chapter_name, chapter_name)
        
        # Добавление заголовка главы
        self.document.add_heading(translated_name, level=1)
        
        # Конвертация контента
        self._convert_markdown_to_docx(content)
        
        # Разрыв страницы после главы
        self.document.add_page_break()

    def compile_diploma(self):
        """Компиляция всего диплома"""
        # Сортировка глав
        chapters = sorted([
            os.path.join(root, file)
            for root, _, files in os.walk(self.chapters_dir)
            for file in files if file == 'content.md'
        ])

        # Обработка каждой главы
        for chapter_path in chapters:
            self._process_chapter(chapter_path)

        # Сохранение документа
        self.document.save(self.output_path)
        print(f"Диплом сохранен в {self.output_path}")

def main():
    diploma_dir = '/home/user/study/diplom/chapters'
    output_path = '/home/user/study/diplom/diploma.docx'
    
    formatter = DiplomaFormatter(diploma_dir, output_path)
    formatter.compile_diploma()

if __name__ == '__main__':
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.styles.style import _ParagraphStyle

class DocumentStyleAnalyzer:
    """
    Класс для глубокого анализа стилей документа Word
    """
    
    def __init__(self, document_path):
        """Инициализация анализатора с путем к документу"""
        self.document_path = document_path
        self.document = Document(document_path)
    
    def analyze_document_styles(self):
        """Полный анализ стилей документа"""
        style_report = {
            "общие_стили": [],
            "параграф_стили": [],
            "символ_стили": [],
            "таблица_стили": [],
            "нумерация_стили": []
        }
        
        # Анализ всех стилей в документе
        for style in self.document.styles:
            try:
                style_info = {
                    "имя": style.name,
                    "тип": str(style.type),
                }
                
                # Безопасная проверка базового стиля
                if hasattr(style, 'base_style') and style.base_style:
                    style_info["базовый_стиль"] = style.base_style.name
                else:
                    style_info["базовый_стиль"] = "Нет"
                
                # Детальный анализ параграф-стилей
                if style.type == 1:  # Paragraph style
                    para_style = style
                    style_info.update({
                        "выравнивание": str(para_style.paragraph_format.alignment) if para_style.paragraph_format.alignment else "Не установлено",
                        "межстрочный_интервал": para_style.paragraph_format.line_spacing if para_style.paragraph_format.line_spacing else "Не установлено",
                        "отступ_первой_строки": str(para_style.paragraph_format.first_line_indent) if para_style.paragraph_format.first_line_indent else "Нет",
                        "интервал_перед": str(para_style.paragraph_format.space_before) if para_style.paragraph_format.space_before else "Нет",
                        "интервал_после": str(para_style.paragraph_format.space_after) if para_style.paragraph_format.space_after else "Нет",
                    })
                    
                    # Параметры шрифта
                    if hasattr(para_style, 'font') and para_style.font:
                        style_info.update({
                            "шрифт": para_style.font.name,
                            "размер_шрифта": str(para_style.font.size),
                            "жирный": para_style.font.bold,
                            "курсив": para_style.font.italic,
                        })
                    
                    style_report["параграф_стили"].append(style_info)
                
                # Для других типов стилей
                elif style.type == 2:  # Character style
                    style_report["символ_стили"].append(style_info)
                elif style.type == 3:  # Table style
                    style_report["таблица_стили"].append(style_info)
                elif style.type == 4:  # Numbering style
                    style_report["нумерация_стили"].append(style_info)
                else:
                    style_report["общие_стили"].append(style_info)
            
            except Exception as e:
                print(f"Ошибка при обработке стиля {style.name}: {e}")
        
        return style_report
    
    def generate_style_report(self):
        """Генерация подробного текстового отчета о стилях"""
        styles = self.analyze_document_styles()
        
        report = "🔍 Полный анализ стилей документа\n\n"
        
        # Параграф стили
        report += "### Стили параграфов:\n"
        for style in styles["параграф_стили"]:
            report += f"#### {style['имя']}\n"
            report += f"- Базовый стиль: {style['базовый_стиль']}\n"
            report += f"- Выравнивание: {style.get('выравнивание', 'Не указано')}\n"
            report += f"- Межстрочный интервал: {style.get('межстрочный_интервал', 'Не указано')}\n"
            report += f"- Отступ первой строки: {style.get('отступ_первой_строки', 'Нет')}\n"
            report += f"- Интервал перед: {style.get('интервал_перед', 'Нет')}\n"
            report += f"- Интервал после: {style.get('интервал_после', 'Нет')}\n"
            
            # Параметры шрифта
            if 'шрифт' in style:
                report += f"- Шрифт: {style['шрифт']}\n"
                report += f"- Размер шрифта: {style.get('размер_шрифта', 'Не указан')}\n"
                report += f"- Жирный: {style.get('жирный', False)}\n"
                report += f"- Курсив: {style.get('курсив', False)}\n"
            
            report += "\n"
        
        # Сохранение отчета
        report_path = os.path.join(os.path.dirname(self.document_path), "document_style_report.md")
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        print(f"📄 Отчет о стилях сохранен в {report_path}")
        return report

def main():
    """Основная функция для запуска анализа стилей"""
    document_path = '/home/user/study/diplom/diploma.docx'
    
    # Проверка существования файла
    if not os.path.exists(document_path):
        print(f"Ошибка: файл {document_path} не найден")
        return
    
    # Создание объекта для анализа стилей
    analyzer = DocumentStyleAnalyzer(document_path)
    analyzer.generate_style_report()

if __name__ == '__main__':
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import subprocess
import time

def run_formatter():
    """Запуск основного форматера диплома"""
    print("Запуск форматирования диплома...")
    formatter_script = "/home/user/study/diplom/scripts/diploma_formatter.py"
    
    try:
        subprocess.run(["python3", formatter_script], check=True)
        print("Основное форматирование завершено успешно")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при форматировании: {e}")
        return False

def run_spacing_fixer():
    """Запуск исправления отступов и интервалов"""
    print("Запуск исправления отступов и интервалов...")
    fixer_script = "/home/user/study/diplom/scripts/document_spacing_fixer.py"
    
    try:
        subprocess.run(["python3", fixer_script], check=True)
        print("Исправление отступов завершено успешно")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при исправлении отступов: {e}")
        return False

def run_validator():
    """Запуск валидатора диплома"""
    print("Запуск валидации диплома...")
    validator_script = "/home/user/study/diplom/scripts/diploma_validator.py"
    
    try:
        subprocess.run(["python3", validator_script], check=True)
        print("Валидация завершена")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при валидации: {e}")
        return False

def main():
    """Основная функция для запуска всего процесса форматирования"""
    print("=== Начало процесса форматирования диплома ===")
    
    # Шаг 1: Основное форматирование
    if not run_formatter():
        print("Процесс остановлен из-за ошибки в основном форматировании")
        return
    
    # Пауза для завершения операций с файлом
    time.sleep(1)
    
    # Шаг 2: Исправление отступов и интервалов
    if not run_spacing_fixer():
        print("Процесс остановлен из-за ошибки в исправлении отступов")
        return
    
    # Пауза для завершения операций с файлом
    time.sleep(1)
    
    # Шаг 3: Валидация результата
    run_validator()
    
    print("=== Процесс форматирования диплома завершен ===")
    print("Результат сохранен в файле: /home/user/study/diplom/diploma.docx")

if __name__ == "__main__":
    main()

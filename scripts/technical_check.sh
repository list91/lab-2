#!/bin/bash

# Функция для проверки шрифта и форматирования в Markdown
check_markdown_formatting() {
    local file="$1"
    local errors=()

    # Проверка использования шрифта
    if grep -qE '(font-family:|font:)' "$file"; then
        errors+=("Обнаружено прямое указание шрифта в файле")
    fi

    # Проверка межстрочного интервала
    if ! grep -qE '^line-height:' "$file"; then
        errors+=("Не указан межстрочный интервал")
    fi

    # Проверка полей
    if ! grep -qE '(margin:|padding:)' "$file"; then
        errors+=("Не указаны отступы")
    fi

    # Вывод ошибок
    if [ ${#errors[@]} -ne 0 ]; then
        echo "Ошибки в файле $file:"
        printf '%s\n' "${errors[@]}"
    fi
}

# Функция для проверки количества символов в строках
check_line_length() {
    local file="$1"
    local max_length=80
    
    while IFS= read -r line; do
        length=${#line}
        if [ $length -gt $max_length ]; then
            echo "Слишком длинная строка в $file: $line (${length} символов)"
        fi
    done < "$file"
}

# Основная функция проверки
check_diploma_technical_aspects() {
    local directory="/home/user/study/diplom/chapters"
    
    echo "Начало технической проверки документов..."
    
    # Поиск всех .md файлов
    find "$directory" -type f -name "*.md" | while read -r file; do
        echo "Проверка файла: $file"
        
        # Проверка форматирования Markdown
        check_markdown_formatting "$file"
        
        # Проверка длины строк
        check_line_length "$file"
    done
}

# Запуск проверки
check_diploma_technical_aspects

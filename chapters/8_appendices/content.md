# 8. Приложения

## Исходный код проекта
### Структура репозитория
tomato_disease_classifier/
├── data/
│   ├── raw/
│   ├── processed/
│   └── test_images/
├── models/
│   ├── svm_classifier.pkl
│   └── scaler.pkl
├── notebooks/
│   ├── data_preprocessing.ipynb
│   └── model_training.ipynb
├── src/
│   ├── preprocessing/
│   │   ├── image_loader.py
│   │   └── data_augmentation.py
│   ├── models/
│   │   ├── svm_model.py
│   │   └── model_evaluation.py
│   └── utils/
│       ├── logging.py
│       └── visualization.py
├── tests/
│   ├── test_preprocessing.py
│   └── test_model.py
├── app/
│   ├── main.py
│   ├── routes.py
│   └── templates/
├── requirements.txt
├── README.md
└── Dockerfile

### Ключевые модули
- `preprocessing`: Подготовка и трансформация изображений
- `models`: Реализация SVM-классификатора
- `utils`: Вспомогательные функции
- `app`: Веб-приложение для классификации

### Инструкции по установке
```bash
# Клонирование репозитория
git clone https://github.com/username/tomato_disease_classifier.git

# Создание виртуального окружения
python3 -m venv venv
source venv/bin/activate

# Установка зависимостей
pip install -r requirements.txt

# Запуск приложения
python app/main.py
```

## Обучающая выборка
### Статистика датасета
- Общее количество изображений: 5000
- Классов болезней: 10
- Изображений на класс: 500
- Разрешение: 64x64 пикселя
- Формат: JPEG, RGB

### Источники данных
- Plant Village Dataset
- Специализированные базы изображений болезней томатов
- Собственные фотографии

### Критерии отбора
- Качество изображений
- Репрезентативность симптомов
- Баланс классов
- Разнообразие условий съемки

## Документация
### Руководство разработчика
#### Настройка окружения
- Требования к системе
- Установка зависимостей
- Конфигурация проекта

#### Архитектура проекта
- Описание компонентов
- Принципы проектирования
- Алгоритмы машинного обучения

#### Развертывание
- Локальный запуск
- Контейнеризация
- Облачное развертывание

### Технические спецификации
- Версии библиотек
- Параметры модели
- Метрики производительности
- Ограничения и допущения

## Руководство пользователя
### Работа с приложением
- Загрузка изображений
- Интерпретация результатов
- Рекомендации по использованию

### Интерфейс
- Описание элементов
- Навигация
- Функциональные возможности

### Сценарии использования
- Диагностика в полевых условиях
- Мониторинг посадок
- Консультации агрономов

## Презентационные материалы
### Слайды
- Краткое описание проекта
- Методология
- Результаты исследования
- Перспективы развития

### Демонстрационное видео
- Работа приложения
- Процесс классификации
- Визуализация результатов

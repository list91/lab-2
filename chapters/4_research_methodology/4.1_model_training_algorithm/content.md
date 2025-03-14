# 4.1 Алгоритм обучения модели

## 1. Подготовка датасета
### Источники данных
- Plant Village Dataset
- Специализированные базы изображений болезней томатов

### Требования к данным
- Разрешение: 256x256 пикселей
- Формат: JPEG
- Цветовой профиль: RGB
- Глубина цвета: 8 бит

### Проверка целостности
- Подсчет общего количества изображений
- Валидация форматов
- Проверка сбалансированности классов

## 2. Загрузка и предобработка изображений
### Чтение изображений
- Использование библиотек OpenCV и Pillow
- Загрузка из директорий классов
- Обработка различных форматов

### Трансформации
- Resize до 64x64 пикселей
- Преобразование в одномерный массив (flatten)
- Нормализация пикселей

### Ограничения
- До 500 изображений на класс
- Случайная выборка при превышении лимита

## 3. Разделение данных
### Стратегия разбиения
- Обучающая выборка: 80%
- Тестовая выборка: 20%
- Фиксированное случайное разделение (random_state=42)

### Техника стратификации
- Сохранение пропорций классов
- Равномерное распределение

## 4. Масштабирование признаков
### StandardScaler
- Центрирование относительно среднего
- Нормализация дисперсии
- Устранение влияния разных шкал признаков

### Параметры масштабирования
- Среднее значение
- Стандартное отклонение
- Сохранение параметров для последующего использования

## 5. Обучение классификатора
### Метод опорных векторов (SVM)
- Ядро: радиальная базисная функция (RBF)
- Включена вероятностная оценка
- Автоматическая балансировка классов

### Настройка гиперпараметров
- Поиск по сетке (Grid Search)
- Кросс-валидация
- Метрики: accuracy, f1-score

## 6. Оценка качества модели
### Метрики классификации
- Accuracy
- Precision
- Recall
- F1-score
- Матрица ошибок

### Визуализация
- ROC-кривая
- Precision-Recall кривая
- Тепловая карта матрицы ошибок

### Логирование
- Запись результатов эксперимента
- Сохранение метрик
- Трекинг версий модели

## 7. Сохранение модели
### Сериализация
- Классификатор SVM
- Параметры масштабирования
- Список классов болезней

### Форматы
- Pickle
- Joblib
- ONNX (для кроссплатформенности)

### Документация
- Версионность модели
- Метаданные эксперимента
- Описание препроцессинга

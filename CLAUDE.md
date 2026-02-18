# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Проект

Исследование продаж на основе данных Excel. Цель - анализ покупателей и групп товаров с экспортом в standalone HTML.

## Структура данных Excel

Файл `test_sales_data.xlsx` содержит:
- **Строка 1**: Годы (2023, 2024, 2025) - объединённые ячейки
- **Строка 2**: Месяцы (Январь-Декабрь) - объединённые ячейки
- **Строка 3**: Группы товаров (28 групп) - объединённые ячейки
- **Строка 4**: Показатели (Количество в чеке, Сумма в чеке, Число чеков) + заголовок "ИД клиента"
- **Строки 5+**: Данные клиентов

Итого: 3 года × 12 месяцев × 28 групп × 3 показателя = 3024 столбца данных

## Команды

```bash
# Активация виртуальной среды
h:\AI_research\10\venv\Scripts\activate

# Запуск скриптов
h:/AI_research/10/venv/Scripts/python.exe <script.py>

# Генерация тестовых данных
h:/AI_research/10/venv/Scripts/python.exe generate_test_data.py

# Установка пакетов
h:/AI_research/10/venv/Scripts/pip.exe install <package>

# Запуск Jupyter Lab
h:/AI_research/10/venv/Scripts/jupyter-lab.exe
```

## Установленные пакеты

- pandas, numpy - работа с данными
- openpyxl - чтение Excel файлов
- jupyter - интерактивные notebook'и
- plotly - интерактивные графики для HTML экспорта

## Основные файлы

- `generate_test_data.py` - генерация тестового Excel файла с 1000 клиентами
- `full_sales_report.ipynb` - **основной notebook** с 50+ графиками и dropdown фильтрами
- `sales_analysis.ipynb` - базовый notebook с исследованием
- `test_sales_data.xlsx` - тестовые данные
- `sales_report.html` - экспортированный интерактивный отчёт

## Структура отчёта (50 графиков)

1. **Общая статистика** (5): KPI, выручка, клиенты по годам, тепловые карты
2. **Анализ покупателей** (20): сегментация по сумме и цене, когортный анализ, миграция, Sankey
3. **Анализ групп товаров** (15): выручка, сезонность, динамика с dropdown фильтрами
4. **Связь покупателей и групп** (10): treemap, sunburst, корреляция, radar профили

## Ключевые особенности кода

- Парсинг сложной шапки Excel через `ffill()` для объединённых ячеек
- Преобразование в long format через `stack()`
- Обязательное преобразование типов: `pd.to_numeric(df_long[col], errors='coerce')`
- Dropdown фильтры через Plotly `updatemenus`

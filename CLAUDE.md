# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Проект

Исследование продаж на основе данных Excel. Цель — анализ покупателей и групп товаров с экспортом в standalone HTML.

## Команды

```bash
# Запуск скриптов
h:/AI_research/10/venv/Scripts/python.exe <script.py>

# Генерация тестовых данных
h:/AI_research/10/venv/Scripts/python.exe generate_test_data.py

# Установка пакетов
h:/AI_research/10/venv/Scripts/pip.exe install <package>

# Запуск Jupyter Lab
h:/AI_research/10/venv/Scripts/jupyter-lab.exe
```

## Структура данных Excel

Файл `test_sales_data.xlsx` — 3 листа (2023, 2024, 2025), по одному на год.

Каждый лист:
- **Строка 1**: Месяцы (Январь–Декабрь) — объединённые ячейки
- **Строка 2**: Группы товаров (28 групп) — объединённые ячейки
- **Строка 3**: Показатели + заголовок "ИД клиента"
- **Строки 4+**: Данные клиентов (1000 клиентов, один `CLIENT_XXXX` на всех листах)

4 показателя в каждой группе: Количество в чеке, Сумма в чеке, Число чеков, Наценка продажи в чеке

Итого на лист: 12 месяцев × 28 групп × 4 показателя = 1344 столбца данных

## Основные файлы

- `generate_test_data.py` — генерация тестового Excel (1000 клиентов, 3 листа)
- `full_sales_report.ipynb` — **основной notebook** с интерактивными графиками и dropdown фильтрами
- `full_sales_report.html` — экспортированный standalone HTML-отчёт
- `test_sales_data.xlsx` — тестовые данные

## Скрипты автоматизации

- `generate_reports.bat` / `.sh` — генерация HTML из notebook'ов через `jupyter nbconvert`
- `git_push.bat` / `.sh` — коммит + push в GitHub

## Установленные пакеты

pandas, numpy, openpyxl, jupyter, plotly

## Ключевые особенности кода

- Парсинг шапки Excel через `ffill()` для объединённых ячеек
- Данные по годам на отдельных листах, связка клиентов по `ИД клиента`
- Преобразование в long format через `stack()`
- Обязательное: `pd.to_numeric(df_long[col], errors='coerce')`
- Dropdown фильтры через Plotly `updatemenus`

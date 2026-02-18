#!/bin/bash

# Генерация отчётов

cd "$(dirname "$0")"

echo "=== Генерация отчётов ==="

echo ""
echo "--- full_sales_report.html ---"
jupyter nbconvert --to html --execute full_sales_report.ipynb --output full_sales_report.html --no-input
if [ $? -ne 0 ]; then
    echo "[Ошибка] Не удалось сгенерировать full_sales_report.html"
fi

echo ""
echo "--- customer_analytics_report.html ---"
jupyter nbconvert --to html --execute customer_analytics.ipynb --output customer_analytics_report.html --no-input
if [ $? -ne 0 ]; then
    echo "[Ошибка] Не удалось сгенерировать customer_analytics_report.html"
fi

echo ""
echo "--- customer_analytics_documentation.docx ---"
python3 generate_documentation.py
if [ $? -ne 0 ]; then
    echo "[Ошибка] Не удалось сгенерировать документацию"
fi

echo ""
echo "=== Готово! ==="

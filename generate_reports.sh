#!/bin/bash

# Генерация отчётов

cd "$(dirname "$0")"

echo "=== Генерация отчётов ==="

echo ""
echo "--- sales_report.html ---"
python3 generate_test_data.py

echo ""
echo "--- customer_analytics_report.html ---"
jupyter nbconvert --to html --execute full_sales_report.ipynb --output full_sales_report.html --no-input 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[Пропущено] Для генерации HTML из notebook используйте Jupyter Lab"
fi

echo ""
echo "=== Готово! ==="

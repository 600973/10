#!/bin/bash

# Генерация отчётов

cd "$(dirname "$0")"

echo "=== Генерация отчётов ==="

echo ""
echo "--- full_sales_report.html ---"
jupyter nbconvert --to notebook --execute full_sales_report.ipynb --output full_sales_report.ipynb --no-input
if [ $? -ne 0 ]; then
    echo "[Ошибка] Не удалось сгенерировать full_sales_report.html"
fi

echo ""
echo "=== Готово! ==="

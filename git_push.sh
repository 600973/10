#!/bin/bash

# Коммит и пуш в GitHub

cd "$(dirname "$0")"

echo "=== Сохранение изменений в GitHub ==="

git add -A

if git diff --cached --quiet; then
    echo "Нет новых изменений для коммита"
else
    git commit -m "Обновление $(date '+%Y-%m-%d %H:%M')"
    echo "Коммит создан"
fi

git push --force origin main
echo "Изменения отправлены в GitHub"

echo ""
echo "=== Готово! ==="

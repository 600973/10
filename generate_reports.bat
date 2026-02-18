@echo off
chcp 65001 >nul

:: Генерация отчётов

cd /d "H:\AI_research\10"

echo === Генерация отчётов ===

echo.
echo --- full_sales_report.html ---
"H:\AI_research\10\venv\Scripts\jupyter.exe" nbconvert --to notebook --execute full_sales_report.ipynb --output full_sales_report.ipynb --no-input
if %errorlevel% neq 0 (
    echo [Ошибка] Не удалось сгенерировать full_sales_report.html
)

echo.
echo === Готово! ===
pause

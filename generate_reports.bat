@echo off
chcp 65001 >nul

:: Генерация отчётов

cd /d "H:\AI_research\10"

echo === Генерация отчётов ===

echo.
echo --- sales_report.html ---
"H:\AI_research\10\venv\Scripts\python.exe" generate_test_data.py

echo.
echo --- customer_analytics_report.html ---
"H:\AI_research\10\venv\Scripts\jupyter-lab.exe" nbconvert --to html --execute full_sales_report.ipynb --output full_sales_report.html --no-input 2>nul
if %errorlevel% neq 0 (
    echo [Пропущено] Для генерации HTML из notebook используйте Jupyter Lab
)

echo.
echo === Готово! ===
pause

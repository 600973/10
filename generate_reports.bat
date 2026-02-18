@echo off
chcp 65001 >nul

:: Генерация отчётов

cd /d "H:\AI_research\10"

echo === Генерация отчётов ===

echo.
echo --- full_sales_report.html ---
"H:\AI_research\10\venv\Scripts\jupyter.exe" nbconvert --to html --execute full_sales_report.ipynb --output full_sales_report.html --no-input
if %errorlevel% neq 0 (
    echo [Ошибка] Не удалось сгенерировать full_sales_report.html
)

echo.
echo --- customer_analytics_report.html ---
"H:\AI_research\10\venv\Scripts\jupyter.exe" nbconvert --to html --execute customer_analytics.ipynb --output customer_analytics_report.html --no-input
if %errorlevel% neq 0 (
    echo [Ошибка] Не удалось сгенерировать customer_analytics_report.html
)

echo.
echo --- customer_analytics_documentation.docx ---
"H:\AI_research\10\venv\Scripts\python.exe" generate_documentation.py
if %errorlevel% neq 0 (
    echo [Ошибка] Не удалось сгенерировать документацию
)

echo.
echo === Готово! ===
pause

@echo off
chcp 65001 >nul

:: Коммит и пуш в GitHub

cd /d "H:\AI_research\10"

echo === Сохранение изменений в GitHub ===

git add -A

git diff --cached --quiet
if %errorlevel% neq 0 (
    for /f "tokens=*" %%d in ('powershell -command "Get-Date -Format 'yyyy-MM-dd HH:mm'"') do set DATETIME=%%d
    git commit -m "Обновление %DATETIME%"
    echo Коммит создан
) else (
    echo Нет новых изменений для коммита
)

git push --force origin main
echo Изменения отправлены в GitHub

echo.
echo === Готово! ===
pause

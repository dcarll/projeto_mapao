@echo off
echo ========================================
echo  DESINSTALACAO - SISTEMA DE LABORATORIOS
echo ========================================
echo.
echo ATENCAO: Isso vai remover:
echo - Bibliotecas Python instaladas
echo - Atalho da area de trabalho
echo.
set /p CONFIRMA="Deseja continuar? (S/N): "

if /i not "%CONFIRMA%"=="S" (
    echo Desinstalacao cancelada.
    pause
    exit /b
)

echo.
echo Removendo bibliotecas...
pip uninstall -y pillow openpyxl

echo.
echo Removendo atalho...
if exist "%USERPROFILE%\Desktop\Sistema Laboratorios.lnk" (
    del "%USERPROFILE%\Desktop\Sistema Laboratorios.lnk"
    echo [OK] Atalho removido
)

echo.
echo ========================================
echo  DESINSTALACAO CONCLUIDA!
echo ========================================
echo.
echo NOTA: A pasta C:\SCHEDULE_LABS e os dados
echo nao foram removidos. Delete manualmente se desejar.
echo.
pause

@echo off
echo ========================================
echo  VERIFICACAO DO SISTEMA
echo ========================================
echo.

echo Verificando Python...
python --version
if errorlevel 1 (
    echo [ERRO] Python nao encontrado
) else (
    echo [OK] Python instalado
)
echo.

echo Verificando bibliotecas...
python -c "import tkinter; print('[OK] tkinter')" 2>nul || echo [ERRO] tkinter
python -c "import PIL; print('[OK] pillow')" 2>nul || echo [ERRO] pillow
python -c "import openpyxl; print('[OK] openpyxl')" 2>nul || echo [ERRO] openpyxl
echo.

echo Verificando arquivos Python...
if exist "main.py" (echo [OK] main.py) else (echo [ERRO] main.py nao encontrado)
if exist "config.py" (echo [OK] config.py) else (echo [ERRO] config.py nao encontrado)
if exist "database.py" (echo [OK] database.py) else (echo [ERRO] database.py nao encontrado)
if exist "gui.py" (echo [OK] gui.py) else (echo [ERRO] gui.py nao encontrado)
if exist "models.py" (echo [OK] models.py) else (echo [ERRO] models.py nao encontrado)
if exist "utils.py" (echo [OK] utils.py) else (echo [ERRO] utils.py nao encontrado)
echo.

echo Verificando pasta compartilhada...
if exist "C:\SCHEDULE_LABS" (
    echo [OK] C:\SCHEDULE_LABS existe
) else (
    echo [AVISO] C:\SCHEDULE_LABS nao encontrada
)
echo.

echo ========================================
echo  VERIFICACAO CONCLUIDA
echo ========================================
pause

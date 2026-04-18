@echo off
setlocal enabledelayedexpansion
title Sistema de Agendamento de Laboratorios

echo ========================================
echo  SISTEMA DE AGENDAMENTO DE LABORATORIOS
echo ========================================
echo.

REM 1. Verifica se Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado no sistema.
    echo Por favor, instale o Python antes de continuar.
    pause
    exit
)

REM 2. Chama a funcao de verificacao de bibliotecas
echo Verificando dependencias...
call :verificar_bibliotecas

echo.
echo Iniciando sistema...
python main.py
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Ocorreu um erro ao executar o sistema.
    pause
)
exit

:verificar_bibliotecas
REM Verificacao rapida se as bibliotecas estao instaladas
REM r=['tksheet','tkcalendar','fpdf','PIL','openpyxl']
python -c "import importlib.util as i; r=['tksheet','tkcalendar','fpdf','PIL','openpyxl']; m=[p for p in r if i.find_spec(p) is None]; exit(1 if m else 0)" >nul 2>&1
if %errorlevel% neq 0 (
    echo [AVISO] Bibliotecas faltando. Instalando requirements.txt...
    python -m pip install -r requirements.txt
) else (
    echo [OK] Todas as bibliotecas ja estao instaladas.
)
goto :eof

@echo off
setlocal enabledelayedexpansion

REM Muda para o diretorio do script
cd /d "%~dp0"

echo ========================================
echo  INSTALACAO - SISTEMA DE LABORATORIOS
echo ========================================
echo.

REM 1. Verifica se Python esta instalado
echo Verificando Python...
python --version
if errorlevel 1 (
    echo.
    echo [ERRO] Python nao encontrado no sistema.
    echo Por favor, instale o Python 3.10 ou superior antes de continuar.
    echo Visite python.org para baixar o instalador.
    echo.
    pause
    exit /b
)
echo [OK] Python detectado!
echo.

REM 2. Upgrade do PIP
echo Atualizando o gerenciador de pacotes PIP...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo [AVISO] Falha ao atualizar o PIP. Tentando continuar...
) else (
    echo [OK] PIP atualizado!
)
echo.

REM 3. Instala dependencias do arquivo requirements.txt
if exist "requirements.txt" (
    echo ========================================
    echo  INSTALANDO BIBLIOTECAS...
    echo ========================================
    echo.
    python -m pip install -r requirements.txt
    echo.
    if errorlevel 1 (
        echo [ERRO] Falha ao instalar dependencias.
        echo Verifique sua conexao com a internet e tente novamente.
        echo.
        pause
        exit /b
    )
    echo [OK] Bibliotecas instaladas com sucesso!
) else (
    echo [ERRO] Arquivo requirements.txt nao encontrado!
    echo Certifique-se de que esta rodando este script na pasta do sistema.
    echo.
    pause
    exit /b
)

echo.
echo ========================================
echo  INSTALACAO CONCLUIDA!
echo ========================================
echo.
echo O sistema esta pronto para ser usado.
echo Inicie o programa atraves de INICIAR.bat
echo.
pause

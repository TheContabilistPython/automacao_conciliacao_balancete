@echo off
title INSTALACAO - Dependencias Python
echo ==============================================
echo    INSTALACAO DAS DEPENDENCIAS PYTHON
echo ==============================================
echo.
echo Verificando se o Python esta instalado...
python --version
if %errorlevel% neq 0 (
    echo.
    echo ERRO: Python nao encontrado!
    echo Por favor, instale o Python 3.7 ou superior de: https://python.org
    echo.
    pause
    exit /b 1
)

echo.
echo Python encontrado! Instalando dependencias...
echo.

pip install --upgrade pip
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo ERRO: Falha na instalacao das dependencias.
    echo Tente executar manualmente: pip install -r requirements.txt
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo SUCCESS: Todas as dependencias foram instaladas com sucesso!
    echo.
)

echo.
echo Pressione qualquer tecla para continuar...
pause >nul

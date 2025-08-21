@echo off
title AUTO BAL V2 - Processamento de Balancete
echo ==============================================
echo    AUTO BAL V2 - Processamento de Balancete
echo ==============================================
echo.
echo Iniciando o script de processamento de balancete...
echo.

cd /d "C:\projeto\robo"

python auto_bal_v2.py

if %errorlevel% neq 0 (
    echo.
    echo ERRO: O script encontrou um problema durante a execucao.
    echo Codigo de erro: %errorlevel%
    echo.
    pause
    exit /b %errorlevel%
) else (
    echo.
    echo SUCCESS: Script executado com sucesso!
    echo.
)

echo.
echo Pressione qualquer tecla para fechar...
pause >nul

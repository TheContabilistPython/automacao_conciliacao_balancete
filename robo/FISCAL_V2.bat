@echo off
title FISCAL V2 - Conciliacao Fiscal
echo ==============================================
echo    FISCAL V2 - Conciliacao Fiscal
echo ==============================================
echo.
echo Iniciando o script de conciliacao fiscal...
echo.

cd /d "C:\projeto\robo"

python FISCAL_V2.py

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

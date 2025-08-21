@echo off
title FOLHA V3 - Conciliacao Folha de Pagamento
echo ==============================================
echo    FOLHA V3 - Conciliacao Folha de Pagamento
echo ==============================================
echo.
echo Iniciando o script de conciliacao de folha...
echo.

cd /d "C:\projeto\robo"

python FOLHA_V3.PY

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

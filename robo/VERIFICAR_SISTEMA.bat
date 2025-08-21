@echo off
title VERIFICACAO DO SISTEMA
color 0A
echo ==============================================
echo    VERIFICACAO DO SISTEMA
echo ==============================================
echo.

echo [1/6] Verificando Python...
python --version
if %errorlevel% neq 0 (
    echo ❌ ERRO: Python nao encontrado!
    goto erro
) else (
    echo ✅ Python instalado corretamente
)

echo.
echo [2/6] Verificando bibliotecas Python...
python -c "import bs4; print('✅ BeautifulSoup4 OK')" 2>nul || echo ❌ BeautifulSoup4 FALTANDO
python -c "import openpyxl; print('✅ OpenPyXL OK')" 2>nul || echo ❌ OpenPyXL FALTANDO
python -c "import pyautogui; print('✅ PyAutoGUI OK')" 2>nul || echo ❌ PyAutoGUI FALTANDO
python -c "import tkinter; print('✅ Tkinter OK')" 2>nul || echo ❌ Tkinter FALTANDO

echo.
echo [3/6] Verificando estrutura de pastas...
if exist "C:\projeto\" (
    echo ✅ Pasta C:\projeto\ existe
) else (
    echo ❌ Pasta C:\projeto\ NAO EXISTE
)

if exist "C:\projeto\empresas.csv" (
    echo ✅ Arquivo empresas.csv existe
) else (
    echo ❌ Arquivo empresas.csv NAO EXISTE
)

if exist "C:\relatorios_fiscal\" (
    echo ✅ Pasta relatorios_fiscal existe
) else (
    echo ❌ Pasta relatorios_fiscal NAO EXISTE
)

if exist "C:\relatorios_folha\" (
    echo ✅ Pasta relatorios_folha existe
) else (
    echo ❌ Pasta relatorios_folha NAO EXISTE
)

echo.
echo [4/6] Verificando scripts Python...
if exist "FISCAL_V2.py" (
    echo ✅ FISCAL_V2.py existe
) else (
    echo ❌ FISCAL_V2.py NAO EXISTE
)

if exist "FOLHA_V3.PY" (
    echo ✅ FOLHA_V3.PY existe
) else (
    echo ❌ FOLHA_V3.PY NAO EXISTE
)

if exist "auto_bal_v2.py" (
    echo ✅ auto_bal_v2.py existe
) else (
    echo ❌ auto_bal_v2.py NAO EXISTE
)

echo.
echo [5/6] Verificando arquivos .bat...
if exist "MENU_PRINCIPAL.bat" (
    echo ✅ MENU_PRINCIPAL.bat existe
) else (
    echo ❌ MENU_PRINCIPAL.bat NAO EXISTE
)

echo.
echo [6/6] Teste final de importacao...
python -c "
try:
    from bs4 import BeautifulSoup
    import openpyxl
    import pyautogui
    import tkinter as tk
    print('✅ TODOS OS MODULOS IMPORTADOS COM SUCESSO!')
    print('🚀 SISTEMA PRONTO PARA USO!')
except Exception as e:
    print('❌ ERRO NA IMPORTACAO:', str(e))
    exit(1)
"

if %errorlevel% neq 0 goto erro

echo.
echo ==============================================
echo    ✅ VERIFICACAO CONCLUIDA COM SUCESSO!
echo    O sistema está pronto para ser usado.
echo ==============================================
goto fim

:erro
echo.
echo ==============================================
echo    ❌ PROBLEMAS ENCONTRADOS!
echo    Consulte o README.md para solucoes.
echo ==============================================

:fim
echo.
echo Pressione qualquer tecla para sair...
pause >nul

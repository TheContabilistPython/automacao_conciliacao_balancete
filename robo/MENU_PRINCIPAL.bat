@echo off
title EXECUTAR TODOS OS SCRIPTS - Menu Principal
color 0A
echo ==============================================
echo    MENU PRINCIPAL - SCRIPTS DE CONCILIACAO
echo ==============================================
echo.
echo Selecione uma opcao:
echo.
echo [1] Executar auto_bal_v2.py (Processamento Balancete)
echo [2] Executar FISCAL_V2.py (Conciliacao Fiscal)
echo [3] Executar FOLHA_V3.py (Conciliacao Folha)
echo [4] Executar TODOS os scripts em sequencia
echo [5] Teste de Extracao do IPI (Diagnostico)
echo [6] Analisar Estrutura HTML (Diagnostico Avancado)
echo [7] Sair
echo.
echo ==============================================

set /p opcao="Digite sua opcao (1-7): "

if "%opcao%"=="1" goto balancete
if "%opcao%"=="2" goto fiscal
if "%opcao%"=="3" goto folha
if "%opcao%"=="4" goto todos
if "%opcao%"=="5" goto teste_ipi
if "%opcao%"=="6" goto analisar_html
if "%opcao%"=="7" goto sair

echo Opcao invalida! Tente novamente.
pause
goto inicio

:balancete
echo.
echo Executando auto_bal_v2.py...
call "C:\projeto\robo\auto_bal_v2.bat"
goto fim

:fiscal
echo.
echo Executando FISCAL_V2.py...
call "C:\projeto\robo\FISCAL_V2.bat"
goto fim

:folha
echo.
echo Executando FOLHA_V3.py...
call "C:\projeto\robo\FOLHA_V3.bat"
goto fim

:teste_ipi
echo.
echo Executando teste de extracao do IPI...
call "C:\projeto\robo\TESTE_IPI.bat"
goto fim

:analisar_html
echo.
echo Executando analisador de estrutura HTML...
call "C:\projeto\robo\ANALISAR_HTML.bat"
goto fim

:todos
echo.
echo ==============================================
echo    EXECUTANDO TODOS OS SCRIPTS
echo ==============================================
echo.
echo 1/3 - Executando Processamento de Balancete...
call "C:\projeto\robo\auto_bal_v2.bat"
echo.
echo 2/3 - Executando Conciliacao Fiscal...
call "C:\projeto\robo\FISCAL_V2.bat"
echo.
echo 3/3 - Executando Conciliacao de Folha...
call "C:\projeto\robo\FOLHA_V3.bat"
echo.
echo ==============================================
echo    TODOS OS SCRIPTS FORAM EXECUTADOS!
echo ==============================================
goto fim

:sair
echo.
echo Saindo...
exit /b 0

:fim
echo.
echo Deseja executar outro script? (S/N)
set /p continuar="Digite S para voltar ao menu ou N para sair: "
if /i "%continuar%"=="S" goto inicio
if /i "%continuar%"=="N" goto sair
echo Opcao invalida! Saindo...
pause
exit /b 0

:inicio
cls
goto inicio

@echo off
title BACKUP PARA MIGRACAO
echo ==============================================
echo    BACKUP PARA MIGRACAO
echo ==============================================
echo.

set DATA=%date:~6,4%-%date:~3,2%-%date:~0,2%
set BACKUP_DIR=C:\BACKUP_CONCILIACAO_%DATA%

echo Criando backup em: %BACKUP_DIR%
mkdir "%BACKUP_DIR%" 2>nul

echo.
echo [1/4] Copiando scripts Python...
copy "FISCAL_V2.py" "%BACKUP_DIR%\" >nul 2>&1
copy "FOLHA_V3.PY" "%BACKUP_DIR%\" >nul 2>&1
copy "auto_bal_v2.py" "%BACKUP_DIR%\" >nul 2>&1

echo [2/4] Copiando scripts .bat...
copy "*.bat" "%BACKUP_DIR%\" >nul 2>&1

echo [3/4] Copiando arquivos de configuracao...
copy "requirements.txt" "%BACKUP_DIR%\" >nul 2>&1
copy "README.md" "%BACKUP_DIR%\" >nul 2>&1
copy "C:\projeto\empresas.csv" "%BACKUP_DIR%\" 2>nul

echo [4/4] Criando estrutura de pastas...
mkdir "%BACKUP_DIR%\estrutura_pastas" >nul 2>&1
echo C:\projeto\robo\ > "%BACKUP_DIR%\estrutura_pastas\pastas_necessarias.txt"
echo C:\projeto\planilhas\ >> "%BACKUP_DIR%\estrutura_pastas\pastas_necessarias.txt"
echo C:\relatorios_fiscal\ >> "%BACKUP_DIR%\estrutura_pastas\pastas_necessarias.txt"
echo C:\relatorios_folha\ >> "%BACKUP_DIR%\estrutura_pastas\pastas_necessarias.txt"

echo.
echo ✅ BACKUP CONCLUIDO!
echo Pasta criada: %BACKUP_DIR%
echo.
echo Para migrar para outra maquina:
echo 1. Copie toda a pasta %BACKUP_DIR% para a nova maquina
echo 2. Execute INSTALAR_DEPENDENCIAS.bat
echo 3. Execute VERIFICAR_SISTEMA.bat
echo 4. Crie a estrutura de pastas conforme README.md
echo.
echo Pressione qualquer tecla para abrir a pasta do backup...
pause >nul
explorer "%BACKUP_DIR%"

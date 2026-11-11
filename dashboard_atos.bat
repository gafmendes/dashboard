@echo off
chcp 65001 > nul
title Dashboard de Atos Judiciais - TJPE

echo ======================================================================
echo    CONSOLIDADOR DE ATOS JUDICIAIS - TJPE (SEGMENTO 8.17)
echo ======================================================================
echo.

REM Verifica se o Python esta instalado e disponivel no PATH
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRO] Python nao foi encontrado no PATH do Windows.
    echo Por favor, instale o Python 3.11+ e marque a opcao "Add Python to PATH".
    pause
    exit /b 1
)

REM Caminho relativo do diretorio deste script
cd /d "%~dp0"

echo [1/2] Verificando dependencias necessarias...
python -m pip install -q -r requirements.txt

echo [2/2] Executando consolidacao incremental...
python main.py %*

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo [SUCESSO] Consolidacao finalizada com exito!
    echo ======================================================================
) else (
    echo.
    echo ======================================================================
    echo [ATENCAO] A execucao encerrou com codigo de erro %errorlevel%.
    echo Verifique o arquivo dashboard_atos.log para mais detalhes.
    echo ======================================================================
)

echo.
echo Pressione qualquer tecla para fechar esta janela...
pause > nul

@echo off
title Planilha Diego - Gerador de Curva ABC e Dashboard
chcp 65001 > nul
cls

echo =======================================================================
echo                     PLANILHA DIEGO - VERSAO 2.0 PRO
echo =======================================================================
echo.
echo  Iniciando o servidor da aplicacao...
echo.

:: Verificar se Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao foi encontrado no sistema!
    echo Por favor, instale o Python 3.10+ para executar o aplicativo.
    pause
    exit /b 1
)

:: Abrir navegador automaticamente apos 2 segundos
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:5000"

:: Executar servidor Flask
python app.py

pause

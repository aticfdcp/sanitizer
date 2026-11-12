@echo off
:: Configurar codificación UTF-8 para ver bien los acentos
chcp 65001 >nul

echo ===============================================
echo      ATIC WORD SANITIZER - LANZADOR
echo ===============================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0ATIC.ps1"

echo.
echo ===============================================
echo Proceso finalizado.
echo ===============================================
pause
@echo off
title 0 - Diagnostico
cd /d "%~dp0"
echo.
echo  Le homologacao e producao, mostra quantos arquivos foram encontrados
echo  em cada situacao e grava um CSV com a lista completa.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0DeployGate.ps1" -Diag -Silencioso
echo.
pause

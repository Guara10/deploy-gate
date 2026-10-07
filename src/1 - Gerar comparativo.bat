@echo off
title 1 - Gerar comparativo
cd /d "%~dp0"
echo.
echo  Comparando homologacao x producao (caminhos em config.psd1).
echo  Na primeira vez pode demorar alguns minutos.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0DeployGate.ps1"
echo.
pause

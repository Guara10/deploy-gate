@echo off
title 3 - Publicar conferidos
cd /d "%~dp0"
echo.
echo  ATENCAO: esta etapa GRAVA NA PRODUCAO.
echo  So rode depois de conferir a pasta separada na etapa 2.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0DeployGate.ps1" -PublicarUltima
echo.
pause

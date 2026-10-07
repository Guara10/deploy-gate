@echo off
title 2 - Separar autorizados
cd /d "%~dp0"
echo.
echo  Procura o decisoes.csv mais recente que voce baixou do relatorio
echo  e COPIA os arquivos autorizados para uma pasta de conferencia.
echo  Nada e movido nem apagado da homologacao.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0DeployGate.ps1" -AplicarUltima
echo.
pause

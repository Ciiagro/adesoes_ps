@echo off
title Gestao da Carreta do Agro (porta 5002)
cd /d "%~dp0"
set SOMENTE_SINDICATOS=0
set SOMENTE_CARRETA=1
echo Abra no navegador: http://localhost:5002/gestao-carreta
python app.py 5002
pause

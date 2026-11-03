@echo off
title Gestao dos Sindicatos Rurais (porta 5001)
cd /d "%~dp0"
set SOMENTE_SINDICATOS=1
set SOMENTE_CARRETA=0
echo Abra no navegador: http://localhost:5001/sindicatos
python app.py 5001
pause

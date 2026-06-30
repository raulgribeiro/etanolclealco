@echo off
title Atualizando Dashboard Gestao Etanol - Clealco
echo.
echo  Instalando dependencias (se necessario)...
pip install selenium openpyxl -q
echo.
echo  Atualizando dashboard...
echo.
cd /d "%~dp0"
python atualizar_dashboard.py

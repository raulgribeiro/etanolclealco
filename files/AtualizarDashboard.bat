@echo off
title Atualizando Dashboard Gestao Etanol - Clealco
echo.
echo  Instalando dependencias (se necessario)...
pip install msal requests openpyxl -q
echo.
echo  Atualizando dashboard...
echo.
cd /d "%~dp0"
python atualizar_dashboard.py

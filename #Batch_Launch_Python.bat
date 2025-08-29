@echo off
chcp 65001 >nul
cd /d "%~dp0"
python #Skapa_elevfoton.py
echo.
echo ======================================================================
echo Skriptet har körts. Tryck valfri tangent för att stänga detta fönster.
echo ======================================================================
pause >nul

@echo off
echo Buscando ofertas... (puede tardar 2-3 minutos)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BuscadorOfertas.ps1" -Forzar

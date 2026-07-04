@echo off
chcp 65001 >nul
set PYTHONUTF8=1
py "%~dp0generer_site.py" || python "%~dp0generer_site.py" || (echo. & echo Python introuvable - installez-le depuis https://www.python.org/downloads/ & echo Pensez a cocher "Add Python to PATH" pendant l'installation. & echo. & pause)

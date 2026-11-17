@echo off
setlocal
cd /d "%~dp0"
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "app.py"
) else (
    echo No se encontro pythonw en el PATH. Se abrira con consola.
    python "app.py"
    pause
)
endlocal

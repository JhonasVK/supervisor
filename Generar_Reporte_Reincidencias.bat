@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
    echo.
    echo No se encontro Node.js instalado en este equipo.
    echo Descargalo desde https://nodejs.org e instalalo, luego vuelve a intentar.
    echo.
    pause
    exit /b 1
)

echo Generando informe de Repetido Reparado...
echo.
node generar_reincidencias.js
if errorlevel 1 (
    echo.
    echo Ocurrio un error generando el informe de Repetido Reparado. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo Generando informe de Averias de Infancia...
echo.
node generar_infancia.js
if errorlevel 1 (
    echo.
    echo Ocurrio un error generando el informe de Averias de Infancia. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo Actualizando indice...
echo.
node generar_indice.js
if errorlevel 1 (
    echo.
    echo Ocurrio un error generando el indice. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo Listo. Se generaron los Excel nuevos y se actualizaron los dashboards.
echo Abre index.html con doble clic para elegir el informe que quieres ver.
pause

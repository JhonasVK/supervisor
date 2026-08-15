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
echo ========================================================
echo   Publicando en GitHub Pages...
echo ========================================================
echo.

git --version >nul 2>&1
if errorlevel 1 (
    echo No se encontro Git instalado. Los informes se generaron localmente
    echo pero no se publicaron. Descarga Git desde https://git-scm.com
    pause
    exit /b 0
)

echo [1/3] Registrando cambios...
git add .
echo       OK

echo [2/3] Guardando version con fecha y hora...
for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HORA=%%a:%%b
git commit -m "Actualizacion %FECHA% %HORA%"

echo [3/3] Subiendo a GitHub...
git push origin master
if errorlevel 1 (
    echo.
    echo No se pudo subir a GitHub. Verifica tu conexion, o si no habia
    echo cambios nuevos que publicar, esto es normal.
    pause
    exit /b 0
)

echo.
echo ========================================================
echo   Listo! Informes actualizados y publicados
echo ========================================================
echo.
echo  Local:  abre index.html con doble clic
echo  Web:    https://jhonasvk.github.io/supervisor/
echo  (la version web tarda ~1 minuto en actualizarse)
echo.
pause

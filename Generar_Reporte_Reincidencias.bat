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
echo Generando Portal de Tecnicos...
echo.
pushd portal-tecnicos
node generar_portal.js
if errorlevel 1 (
    echo.
    echo Ocurrio un error generando el Portal de Tecnicos. Revisa el mensaje de arriba.
    popd
    pause
    exit /b 1
)
popd

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

for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HORA=%%a:%%b

where gh >nul 2>nul
if not errorlevel 1 (
    gh auth switch --hostname github.com --user JhonasVK >nul 2>nul
)

echo [Supervisor] Registrando cambios...
git add .
git commit -m "Actualizacion %FECHA% %HORA%"

echo [Supervisor] Subiendo a GitHub...
git push origin master
if errorlevel 1 (
    echo.
    echo No se pudo subir "Supervisor" a GitHub. Verifica tu conexion, o si no habia
    echo cambios nuevos que publicar, esto es normal.
)

echo.
echo [Portal Tecnicos] Registrando cambios...
pushd portal-tecnicos

where gh >nul 2>nul
if not errorlevel 1 (
    gh auth switch --hostname github.com --user supervisionenaccion-stack >nul 2>nul
)

git add .
git commit -m "Actualizacion %FECHA% %HORA%"

echo [Portal Tecnicos] Subiendo a GitHub...
git push origin master
if errorlevel 1 (
    echo.
    echo No se pudo subir "Portal Tecnicos" a GitHub. Verifica tu conexion, o si no habia
    echo cambios nuevos que publicar, esto es normal.
)
popd

echo.
echo ========================================================
echo   Listo! Informes actualizados y publicados
echo ========================================================
echo.
echo  Supervisor:      https://jhonasvk.github.io/supervisor/
echo  Portal Tecnicos: https://supervisionenaccion-stack.github.io/portal-tecnicos/
echo  (la version web tarda ~1 minuto en actualizarse)
echo.
pause

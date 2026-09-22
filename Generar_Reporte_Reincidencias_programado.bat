@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
    echo No se encontro Node.js instalado en este equipo.
    msg "%USERNAME%" "Informes COBRA: no se encontro Node.js instalado."
    exit /b 1
)

echo Generando informe de Repetido Reparado...
node generar_reincidencias.js
if errorlevel 1 (
    echo Ocurrio un error generando el informe de Repetido Reparado.
    msg "%USERNAME%" "Informes COBRA: fallo generando Repetido Reparado. Revisa log_reportes.txt"
    exit /b 1
)

echo.
echo Generando informe de Averias de Infancia...
node generar_infancia.js
if errorlevel 1 (
    echo Ocurrio un error generando el informe de Averias de Infancia.
    msg "%USERNAME%" "Informes COBRA: fallo generando Averias de Infancia. Revisa log_reportes.txt"
    exit /b 1
)

echo.
echo Generando resumen para correo...
if exist Resumen_Diario_Correo.png del Resumen_Diario_Correo.png
node generar_resumen_correo.js
if exist Resumen_Diario_Correo.png (
    echo ==^> OK: Resumen_Diario_Correo.png generado correctamente.
) else (
    echo ==^> AVISO: no se genero Resumen_Diario_Correo.png.
)

echo.
echo Copiando INF-09 (produccion / productividad) desde OneDrive...
call copiar_baremos.bat

echo.
echo Generando informe de Produccion por tecnicos...
node generar_produccion.js
if errorlevel 1 (
    echo Ocurrio un error generando el informe de Produccion.
    msg "%USERNAME%" "Informes COBRA: fallo generando Produccion. Revisa log_reportes.txt"
    exit /b 1
)

echo.
echo Actualizando indice...
node generar_indice.js
if errorlevel 1 (
    echo Ocurrio un error generando el indice.
    msg "%USERNAME%" "Informes COBRA: fallo generando el indice. Revisa log_reportes.txt"
    exit /b 1
)

echo.
echo Generando Portal de Tecnicos...
pushd portal-tecnicos
node generar_portal.js
if errorlevel 1 (
    echo Ocurrio un error generando el Portal de Tecnicos.
    popd
    msg "%USERNAME%" "Informes COBRA: fallo generando el Portal de Tecnicos. Revisa log_reportes.txt"
    exit /b 1
)
popd

echo.
echo ========================================================
echo   Publicando en GitHub Pages...
echo ========================================================

git --version >nul 2>&1
if errorlevel 1 (
    echo No se encontro Git instalado. Los informes se generaron localmente pero no se publicaron.
    msg "%USERNAME%" "Informes COBRA: se generaron pero no se publicaron (falta Git)."
    exit /b 0
)

for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HORA=%%a:%%b

set FALLO_SUPERVISOR=0
set FALLO_PORTAL=0

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
    echo No se pudo subir "Supervisor" a GitHub.
    set FALLO_SUPERVISOR=1
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
    echo No se pudo subir "Portal Tecnicos" a GitHub.
    set FALLO_PORTAL=1
)
popd

echo.
if "%FALLO_SUPERVISOR%%FALLO_PORTAL%"=="00" (
    echo Listo! Informes actualizados y publicados en los dos repos.
    msg "%USERNAME%" "Informes COBRA: se actualizaron y publicaron correctamente (Supervisor + Portal Tecnicos)."
) else (
    msg "%USERNAME%" "Informes COBRA: fallo al publicar (Supervisor=%FALLO_SUPERVISOR% Portal=%FALLO_PORTAL%). Revisa log_reportes.txt"
)

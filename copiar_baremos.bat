@echo off
setlocal
REM Copia los archivos INF de cada mes (INF-09 Septiembre, INF-10 Octubre...;
REM produccion / baremos Punta Arenas + Coyhaique)
REM desde el OneDrive local a la carpeta baremosTigo, para que generar_portal.js
REM pueda calcular la productividad de cada tecnico.
REM
REM Si Francisco sube una version nueva a la carpeta compartida "JVK", OneDrive
REM la sincroniza sola y este .bat trae la copia mas reciente.

set "ORIGEN=%USERPROFILE%\OneDrive - COBRA CHILE SERVICIOS S.A\Archivos de Francisco Flores Villegas - JVK"
set "DESTINO=%~dp0baremosTigo"

if not exist "%DESTINO%" mkdir "%DESTINO%"

if not exist "%ORIGEN%\INF-*.xlsx" (
    echo AVISO: no se encontro ningun INF-*.xlsx en:
    echo   "%ORIGEN%"
    echo Revisa que la carpeta compartida "JVK" siga con acceso directo en tu OneDrive.
    exit /b 0
)

REM robocopy reintenta si Excel tiene el archivo abierto un momento (/R:3 /W:3).
robocopy "%ORIGEN%" "%DESTINO%" "INF-*.xlsx" /R:3 /W:3 /NJH /NJS /NDL /NP /NC /NS >nul
if %ERRORLEVEL% GEQ 8 (
    echo AVISO: no se pudo copiar el INF a baremosTigo (archivo en uso o sin acceso^).
    echo El portal usara la ultima copia disponible.
    exit /b 0
)

echo ==^> INF copiados/actualizados en baremosTigo.
exit /b 0

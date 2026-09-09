@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   Actualizando datos del panel (DS HNOS)
echo ============================================

rem --- Ubicar node ---
where node >nul 2>&1
if %errorlevel%==0 (
    set "NODE_CMD=node"
) else if exist "C:\Program Files\nodejs\node.exe" (
    set "NODE_CMD=C:\Program Files\nodejs\node.exe"
) else if exist "%LOCALAPPDATA%\Programs\nodejs\node.exe" (
    set "NODE_CMD=%LOCALAPPDATA%\Programs\nodejs\node.exe"
) else (
    echo ERROR: no se encontro "node" ^(ni en PATH ni en las ubicaciones habituales^).
    pause
    exit /b 1
)

rem --- Archivo Excel: el que se arrastro sobre este .bat, o el mas nuevo en esta carpeta ---
set "EXCEL_FILE=%~1"
if "%EXCEL_FILE%"=="" (
    for /f "delims=" %%f in ('dir /b /o-d "*.xlsm" 2^>nul') do (
        if "%EXCEL_FILE%"=="" set "EXCEL_FILE=%%f"
    )
)
if "%EXCEL_FILE%"=="" (
    echo ERROR: no encontre ningun archivo .xlsm.
    echo Arrastra el Excel actualizado sobre este .bat, o dejalo en esta carpeta.
    pause
    exit /b 1
)

echo.
echo Usando archivo: %EXCEL_FILE%
echo.

rem --- Instalar dependencias la primera vez (si falta "xlsx") ---
if not exist "node_modules\xlsx" (
    if not exist "..\node_modules\xlsx" (
        echo Primera vez: instalando dependencias ^(npm install^)...
        call npm install
        if errorlevel 1 (
            echo ERROR: fallo "npm install". Revisa tu conexion a internet.
            pause
            exit /b 1
        )
    )
)

"%NODE_CMD%" importar_datos.js "%EXCEL_FILE%"
if errorlevel 1 (
    echo.
    echo ERROR: fallo la actualizacion. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
pause

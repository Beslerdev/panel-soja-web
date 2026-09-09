@echo off
cd /d "%~dp0"
echo Instalando dependencias (solo tarda la primera vez)...
call npm install
echo.
echo Iniciando el panel...
echo Cuando diga "escuchando en el puerto 3000", abri http://localhost:3000 en el navegador.
echo Para cortar el servidor, cerra esta ventana o apreta Ctrl+C.
echo.
call npm start
pause

@echo off
title Actualizar Portal HME - Git Auto Sync
color 0a

echo ========================================================
echo   ACTUALIZANDO DATOS DEL SISTEMA HME EN EL PORTAL WEB
echo ========================================================
echo.

cd /d "D:\PROYECTOS\CONTROL_ACTIVOS_HME"

echo [1/4] Agregando los cambios del Excel a Git...
git add "DRIVE TRHU BASE.xlsx"

echo [2/4] Creando el paquete de actualizacion...
git commit -m "Actualizacion automatica de datos Excel - %date% %time%"

echo [3/4] Subiendo los cambios a GitHub (Nube)...
git push origin principal

echo.
echo ========================================================
echo   ¡PROCESO COMPLETADO! Tu portal web se actualizara solo.
echo ========================================================
pause
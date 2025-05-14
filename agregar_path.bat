@echo off
setlocal EnableDelayedExpansion

echo --------------------------------------
echo 🛠️ Configurando entorno del sistema...
echo --------------------------------------

:: Inicializar variables vacías
set "LIBREOFFICE_PATH="
set "GHOSTSCRIPT_PATH="
set "PYTHON_PATH="
set "PYTHON_SCRIPTS="

:: 🔍 Detectar LibreOffice
set "L_PATH=C:\Program Files\LibreOffice\program"
if exist "%L_PATH%\soffice.exe" (
    set "LIBREOFFICE_PATH=%L_PATH%"
    echo ✅ LibreOffice detectado en: %LIBREOFFICE_PATH%
) else (
    echo ❌ LibreOffice no encontrado. No se agregará.
)

:: 🔍 Detectar Ghostscript automáticamente
for /D %%G in ("C:\Program Files\gs\gs*") do (
    if exist "%%G\bin\gswin64c.exe" (
        set "GHOSTSCRIPT_PATH=%%G\bin"
    )
)

if defined GHOSTSCRIPT_PATH (
    echo ✅ Ghostscript detectado en: %GHOSTSCRIPT_PATH%
) else (
    echo ❌ Ghostscript no encontrado. No se agregará.
)

:: 🔍 Detectar Python (v3.11 o v3.13)
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    set "PYTHON_PATH=%LocalAppData%\Programs\Python\Python311"
)
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    set "PYTHON_PATH=%LocalAppData%\Programs\Python\Python313"
)

if defined PYTHON_PATH (
    set "PYTHON_SCRIPTS=%PYTHON_PATH%\Scripts"
    echo ✅ Python detectado en: %PYTHON_PATH%
) else (
    echo ❌ Python no encontrado. No se agregará.
)

:: Leer PATH actual desde el registro
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "OLD_PATH=%%b"

echo.
echo 🔍 Verificando y agregando rutas necesarias al PATH...

:: Agregar cada ruta detectada (si fue definida)
if defined LIBREOFFICE_PATH call :agregar_al_path "%LIBREOFFICE_PATH%"
if defined GHOSTSCRIPT_PATH call :agregar_al_path "%GHOSTSCRIPT_PATH%"
if defined PYTHON_PATH call :agregar_al_path "%PYTHON_PATH%"
if defined PYTHON_SCRIPTS call :agregar_al_path "%PYTHON_SCRIPTS%"

echo.
echo ✅ Proceso completado. Reinicia la PC o terminal para aplicar los cambios.
pause
exit /b

:agregar_al_path
set "RUTA=%~1"
echo %OLD_PATH% | find /I "%RUTA%" >nul
if errorlevel 1 (
    echo ➕ Agregando al PATH: %RUTA%
    setx PATH "%OLD_PATH%;%RUTA%" /M
) else (
    echo ⚠️ Ya estaba en el PATH: %RUTA%
)
exit /b

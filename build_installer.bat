@echo off
setlocal

cd /d %~dp0

:: Clean previous builds
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

set PYINSTALLER_PATH=".\venv\Scripts\pyinstaller.exe"

echo ---------------------------------------------------
echo Step 1: Compiling Uploader Service (PyWin32)...
echo ---------------------------------------------------

:: --hidden-import win32timezone is CRITICAL for pywin32 services to run
call %PYINSTALLER_PATH% --noconfirm --onefile --windowed --hidden-import win32timezone --name "uploader_service" uploader_service.py

if not exist "dist\uploader_service.exe" (
    echo Compilation of service failed.
    pause
    exit /b 1
)

echo.
echo ---------------------------------------------------
echo Step 2: Compiling Installer...
echo ---------------------------------------------------
:: We now only bundle uploader_service.exe, no nssm.
call %PYINSTALLER_PATH% --noconfirm --onefile --windowed ^
    --name "SharePointUploaderSetup" ^
    --add-binary "dist/uploader_service.exe;." ^
    installer_gui.py

if exist "dist\SharePointUploaderSetup.exe" (
    echo.
    echo ===================================================
    echo SUCCESS!
    echo Installer located at: dist\SharePointUploaderSetup.exe
    echo ===================================================
) else (
    echo Failed to create installer.
)

pause
endlocal
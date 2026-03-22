@echo off
setlocal

REM ============================================================
REM  PPT Toolbox — No-Installer (xcopy) Deployment
REM  Use this when you cannot run an .exe installer on a machine.
REM  Run this script as the TARGET USER (not admin).
REM ============================================================

set SRCDIR=%~dp0..\PPTToolbox\bin\Release
set DESTDIR=%APPDATA%\PPTToolbox

echo.
echo  ===================================
echo   PPT Toolbox — xcopy Deploy
echo  ===================================
echo.

REM Copy files
echo Copying files to %DESTDIR% ...
if not exist "%DESTDIR%" mkdir "%DESTDIR%"
xcopy /E /I /Y "%SRCDIR%\*" "%DESTDIR%\" >nul
if %errorlevel% neq 0 (
    echo ERROR: File copy failed.
    pause
    exit /b 1
)

REM Register add-in in HKCU (no admin needed)
echo Registering add-in with PowerPoint...
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox" /v "Description"  /t REG_SZ /d "PPT Toolbox" /f >nul
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox" /v "FriendlyName" /t REG_SZ /d "PPT Tools"   /f >nul
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTToolbox" /v "Manifest"     /t REG_SZ /d "%DESTDIR%\PPTToolbox.vsto|vstolocal" /f >nul

echo.
echo  Done! Open PowerPoint to see the "PPT Tools" ribbon tab.
echo.
pause

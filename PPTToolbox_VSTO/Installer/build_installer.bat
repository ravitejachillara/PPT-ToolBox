@echo off
setlocal

REM ============================================================
REM  PPT Toolbox — One-click Installer Builder
REM  Requirements:
REM    1. Inno Setup 6  (default install path used below)
REM    2. The Release build already compiled in VS
REM ============================================================

set INNO="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
set SCRIPT=%~dp0setup.iss
set DISTDIR=%~dp0..\dist

echo.
echo  ===========================
echo   PPT Toolbox Installer Build
echo  ===========================
echo.

REM Check Inno Setup is installed
if not exist %INNO% (
    echo ERROR: Inno Setup 6 not found at %INNO%
    echo Download from: https://jrsoftware.org/isinfo.php
    pause
    exit /b 1
)

REM Build the Release configuration first (requires MSBuild / VS in PATH)
where msbuild >nul 2>&1
if %errorlevel% == 0 (
    echo [1/2] Building Release...
    msbuild "%~dp0..\PPTToolbox\PPTToolbox.csproj" /p:Configuration=Release /t:Build /nologo /v:minimal
    if %errorlevel% neq 0 (
        echo ERROR: MSBuild failed. Fix build errors first.
        pause
        exit /b 1
    )
) else (
    echo [1/2] MSBuild not in PATH — skipping build step.
    echo       Make sure Release binaries already exist in bin\Release\
)

REM Create dist folder
if not exist "%DISTDIR%" mkdir "%DISTDIR%"

REM Compile installer
echo [2/2] Compiling installer...
%INNO% "%SCRIPT%"

if %errorlevel% == 0 (
    echo.
    echo  SUCCESS: PPTToolbox_Setup.exe created in dist\
    echo.
    explorer "%DISTDIR%"
) else (
    echo ERROR: Inno Setup compilation failed.
)

pause

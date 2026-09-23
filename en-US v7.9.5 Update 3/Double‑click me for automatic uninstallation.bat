@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
title CST Slow-Wave Structure User Watch MacroKit - One-Click Uninstallation

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ Info ] Administrator privileges required, restarting with elevation...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo ==============================================
echo     CST Slow-Wave Structure User Watch MacroKit
echo     One-Click Uninstallation
echo ==============================================
echo.

set "TMPFILE=%~dp0_cst_detect.tmp"
powershell -NoProfile -Command "Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*CST Studio Suite*' -and $_.DisplayName -match '202[4-9]' } | Sort-Object DisplayName -Descending | ForEach-Object { if ($_.InstallLocation) { $_.InstallLocation.TrimEnd('\') } elseif ($_.DisplayIcon) { (Split-Path $_.DisplayIcon).TrimEnd('\') } } | Select-Object -Unique | Out-File -FilePath \"%~dp0_cst_detect.tmp\" -Encoding ascii"

set "CSTDIR="
set "idx=0"
set "PREV="
echo Detecting CST installation path...
echo.
for /f "usebackq delims=" %%p in ("%TMPFILE%") do (
    if not "%%p"=="!PREV!" (
        set /a idx+=1
        set "VER!idx!=%%p"
        echo   [!idx!] %%p
        set "PREV=%%p"
    )
)
echo.
del "%TMPFILE%" >nul 2>&1

if !idx!==0 (
    echo [ Info ] No CST Studio Suite 2024 or later detected.
    echo          Continuing with the user macro directory only.
    echo.
    goto :clean_macros
)

echo Detected !idx! CST version(s). Uninstalling from all of them.
echo.

set /a k=1
:uninstall_loop
if !k! gtr !idx! goto :clean_macros

set "CSTDIR="
for %%i in (!k!) do set "CSTDIR=!VER%%i!"
if defined CSTDIR if "!CSTDIR:~-1!"=="\" set "CSTDIR=!CSTDIR:~0,-1!"

set "DST_DLL=!CSTDIR!\AMD64"

echo ---------- [ !k!/!idx! ] !CSTDIR! ----------

if exist "!DST_DLL!\LinXi.dll" (
    del /q "!DST_DLL!\LinXi.dll"
    echo   [ OK ] Removed LinXi.dll
) else (
    echo   [ Info ] LinXi.dll not found, skipped.
)

if exist "!DST_DLL!\LinXi.ini" (
    del /q "!DST_DLL!\LinXi.ini"
    echo   [ OK ] Removed LinXi.ini
) else (
    echo   [ Info ] LinXi.ini not found, skipped.
)

if exist "!DST_DLL!\LinXi_Watch.bas" (
    del /q "!DST_DLL!\LinXi_Watch.bas"
    echo   [ OK ] Removed LinXi_Watch.bas
) else (
    echo   [ Info ] LinXi_Watch.bas not found, skipped.
)

if exist "!DST_MCR!\Check LinXi Version.mcr" (
    echo   [ Info ] Check LinXi Version.mcr is removed together with the macro directory.
)

echo.
set /a k+=1
goto :uninstall_loop

:clean_macros
set "DST_MACROS=%APPDATA%\Dassault Systemes\CST STUDIO SUITE\Library\Macros"
set "DST_MCR=!DST_MACROS!\Admin LinXi Macro"

echo ---------- User macro directory ----------
if exist "!DST_MCR!" (
    rd /s /q "!DST_MCR!"
    if exist "!DST_MCR!" (
        echo   [ Error ] Failed to remove: !DST_MCR!
        echo             The directory may be locked by CST. Close CST and retry.
    ) else (
        echo   [ OK ] Removed the whole macro directory: !DST_MCR!
    )
) else (
    echo   [ Info ] !DST_MCR! not found, skipped.
)
echo.

echo [ Done ] Uninstallation completed successfully!
echo.
pause

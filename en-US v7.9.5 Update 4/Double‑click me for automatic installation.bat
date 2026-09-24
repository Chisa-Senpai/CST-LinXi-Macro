@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
title CST Slow-Wave Structure User Watch MacroKit - One-Click Deployment

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ Info ] Administrator privileges required, restarting with elevation...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo ==============================================
echo     CST Slow-Wave Structure User Watch MacroKit
echo     One-Click Deployment
echo ==============================================
echo.

set "SRC_LIB=%~dp0MacroKit\LinXi.dll"
set "SRC_VB1=%~dp0MacroKit\LinXi.bas"
set "SRC_VB2=%~dp0MacroKit\Me.bas"
set "SRC_VB3=%~dp0MacroKit\She.bas"
set "SRC_INI=%~dp0MacroKit\LinXi.ini"

if not exist "%SRC_LIB%" (
    echo [ Error ] Source file not found: %SRC_LIB%
    echo           Please make sure "MacroKit\LinXi.dll" exists in the script directory.
    pause
    exit /b 1
)
if not exist "%SRC_VB1%" (
    echo [ Error ] Source file not found: %SRC_VB1%
    echo           Please make sure "MacroKit\LinXi.bas" exists in the script directory.
    pause
    exit /b 1
)
if not exist "%SRC_VB2%" (
    echo [ Error ] Source file not found: %SRC_VB2%
    echo           Please make sure "MacroKit\Me.bas" exists in the script directory.
    pause
    exit /b 1
)
if not exist "%SRC_VB3%" (
    echo [ Error ] Source file not found: %SRC_VB3%
    echo           Please make sure "MacroKit\She.bas" exists in the script directory.
    pause
    exit /b 1
)
if not exist "%SRC_INI%" (
    echo [ Error ] Source file not found: %SRC_INI%
    echo           Please make sure "MacroKit\LinXi.ini" exists in the script directory.
    pause
    exit /b 1
)

echo [ OK ] Source file check passed:
echo      %SRC_LIB%
echo      %SRC_VB1%
echo      %SRC_VB2%
echo      %SRC_VB3%
echo      %SRC_INI%
echo.

set "TMPFILE=%~dp0_cst_detect.tmp"
powershell -NoProfile -Command "& { $items = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*CST Studio Suite*' }; $ok = $items | Where-Object { $m = [regex]::Match($_.DisplayName, 'CST Studio Suite[^\d]*(\d{4})'); $m.Success -and ([int]$m.Groups[1].Value -ge 2024) }; $ok | Sort-Object DisplayName -Descending | ForEach-Object { if ($_.InstallLocation) { $_.InstallLocation.TrimEnd('\') } elseif ($_.DisplayIcon) { (Split-Path $_.DisplayIcon).TrimEnd('\') } } | Select-Object -Unique; if ($ok.Count -eq 0 -and $items.Count -gt 0) { '__LOW_VERSION__' } } | Out-File -FilePath \"%~dp0_cst_detect.tmp\" -Encoding ascii"

set "CSTDIR="
set "idx=0"
set "PREV="
echo Detecting CST installation path...
echo.

set "HASLOW=0"
for /f "usebackq delims=" %%p in ("%TMPFILE%") do (
    if "%%p"=="__LOW_VERSION__" (
        set "HASLOW=1"
    ) else (
        if not "%%p"=="!PREV!" (
            set /a idx+=1
            set "VER!idx!=%%p"
            echo   [!idx!] %%p
            set "PREV=%%p"
        )
    )
)

echo.
del "%TMPFILE%" >nul 2>&1

if !idx!==0 (
    if "!HASLOW!"=="1" (
        echo [ Info ] CST Studio Suite detected, but the version is below 2024.
        echo          This tool only supports CST 2024 / 2025 / 2026. Installation skipped.
    ) else (
        echo [ Info ] CST Studio Suite not detected.
        echo          Please install CST 2024 or later and run this script again. Installation skipped.
    )
    pause
    exit /b 1
)

echo Detected !idx! CST version(s), all 2024 or later. Installing MacroKit into all of them, no selection needed.
echo.

set "DST_MACROS=%APPDATA%\Dassault Systemes\CST STUDIO SUITE\Library\Macros"
set "DST_MCR=!DST_MACROS!\Admin LinXi Macro"

echo ---------- User macro directory ----------
echo      !DST_MCR!
echo.

if not exist "!DST_MCR!" mkdir "!DST_MCR!"

copy /y "%SRC_VB1%" "!DST_MCR!\Run LinXi Macro.mcr" >nul
if errorlevel 1 (
    echo [ Error ] Failed to copy LinXi.bas.
    pause
    exit /b 1
)
echo [ OK ] LinXi.bas deployed to:
echo      !DST_MCR!\Run LinXi Macro.mcr
echo.

copy /y "%SRC_VB2%" "!DST_MCR!\Define LinXi Macro.mcr" >nul
if errorlevel 1 (
    echo [ Error ] Failed to copy Me.bas.
    pause
    exit /b 1
)
echo [ OK ] Me.bas deployed to:
echo      !DST_MCR!\Define LinXi Macro.mcr
echo.

copy /y "%SRC_VB3%" "!DST_MCR!\Check LinXi Version.mcr" >nul
if errorlevel 1 (
    echo [ Error ] Failed to copy She.bas.
    pause
    exit /b 1
)
echo [ OK ] She.bas deployed to:
echo      !DST_MCR!\Check LinXi Version.mcr
echo.

set /a k=1

:install_loop
if !k! gtr !idx! goto :install_done

set "CSTDIR="
for %%i in (!k!) do set "CSTDIR=!VER%%i!"
if defined CSTDIR if "!CSTDIR:~-1!"=="\" set "CSTDIR=!CSTDIR:~0,-1!"

echo ============================================================
echo ---------- [ !k!/!idx! ] CST installation root: !CSTDIR! ----------
echo.

set "DST_DLL=!CSTDIR!\AMD64"

echo Target deployment directory:
echo     [ LinXi.dll       ]  !DST_DLL!
echo     [ LinXi.ini       ]  !DST_DLL!
echo     [ LinXi_Watch.bas ]  !DST_DLL!
echo.

if not exist "!DST_DLL!" mkdir "!DST_DLL!"

copy /y "%SRC_LIB%" "!DST_DLL!\LinXi.dll" >nul
if errorlevel 1 (
    echo [ Error ] Failed to copy LinXi.dll.
    pause
    exit /b 1
)
echo [ OK ] LinXi.dll deployed to:
echo      !DST_DLL!\LinXi.dll
echo.

copy /y "%SRC_INI%" "!DST_DLL!\LinXi.ini" >nul
if errorlevel 1 (
    echo [ Error ] Failed to copy LinXi.ini.
    pause
    exit /b 1
)
echo [ OK ] LinXi.ini deployed to:
echo      !DST_DLL!\LinXi.ini
echo.

copy /y "%SRC_VB1%" "!DST_DLL!\LinXi_Watch.bas" >nul
if errorlevel 1 (
    echo [ Error ] Failed to copy LinXi_Watch.bas.
    pause
    exit /b 1
)
echo [ OK ] LinXi.bas deployed as the watch macro:
echo      !DST_DLL!\LinXi_Watch.bas
echo.

set /a k+=1
goto :install_loop

:install_done
echo ============================================================
echo [ Done ] Deployment completed successfully - processed !idx! CST version(s), no manual selection required.
echo          To uninstall, run the uninstaller in the same directory.
echo ============================================================
echo.
pause
exit /b 0

@echo off
setlocal EnableDelayedExpansion
chcp 936 >nul
cd /d "%~dp0"
title CST 慢波结构用户监视器 MacroKit 一键部署

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ 提示 ] 需要管理员权限，正在重新启动...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo ==============================================
echo     CST 慢波结构用户监视器 MacroKit 一键部署
echo ==============================================
echo.

set "SRC_LIB=%~dp0MacroKit\LinXi.dll"
set "SRC_VB1=%~dp0MacroKit\LinXi.bas"
set "SRC_VB2=%~dp0MacroKit\Me.bas"
set "SRC_VB3=%~dp0MacroKit\She.bas"
set "SRC_INI=%~dp0MacroKit\LinXi.ini"

if not exist "%SRC_LIB%" (
    echo [ 错误 ] 未找到源文件: %SRC_LIB%
    echo        请确认脚本目录下有 "MacroKit\LinXi.dll"
    pause
    exit /b 1
)
if not exist "%SRC_VB1%" (
    echo [ 错误 ] 未找到源文件: %SRC_VB1%
    echo        请确认脚本目录下有 "MacroKit\LinXi.bas"
    pause
    exit /b 1
)
if not exist "%SRC_VB2%" (
    echo [ 错误 ] 未找到源文件: %SRC_VB2%
    echo        请确认脚本目录下有 "MacroKit\Me.bas"
    pause
    exit /b 1
)
if not exist "%SRC_VB3%" (
    echo [ 错误 ] 未找到源文件: %SRC_VB3%
    echo        请确认脚本目录下有 "MacroKit\She.bas"
    pause
    exit /b 1
)
if not exist "%SRC_INI%" (
    echo [ 错误 ] 未找到源文件: %SRC_INI%
    echo        请确认脚本目录下有 "MacroKit\LinXi.ini"
    pause
    exit /b 1
)

echo [ OK ] 源文件检查通过:
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
echo 正在检测 CST 安装路径...
echo.

set "HASLOW=0"
for /f "usebackq delims=" %%p in ("%TMPFILE%") do (
    if "%%p"=="__LOW_VERSION__" (
        set "HASLOW=1"
    ) else (
        if not "%%p"=="!PREV!" (
            set /a idx+=1
            set "VER!idx!=%%p"
            echo   [ !idx! ] %%p
            set "PREV=%%p"
        )
    )
)

echo.
del "%TMPFILE%" >nul 2>&1

if !idx!==0 (
    if "!HASLOW!"=="1" (
        echo [ 提示 ] 检测到 CST Studio Suite，但版本低于 2024。
        echo         本程序仅支持 CST 2024 / 2025 / 2026，安装已跳过。
    ) else (
        echo [ 提示 ] 未检测到 CST Studio Suite。
        echo         请先安装 CST 2024 及以上版本后再运行，安装已跳过。
    )
    pause
    exit /b 1
)

echo 检测到 !idx! 个 CST 版本 (均为 2024 及以上), 将全部安装 MacroKit, 无需选择。
echo.

rem ---- 用户宏目录与 CST 版本无关, 只部署一次 ----
set "DST_MACROS=%APPDATA%\Dassault Systemes\CST STUDIO SUITE\Library\Macros"
set "DST_MCR=!DST_MACROS!\Admin LinXi Macro"

echo ---------- 用户宏目录 ----------
echo      !DST_MCR!
echo.

if not exist "!DST_MCR!" mkdir "!DST_MCR!"

copy /y "%SRC_VB1%" "!DST_MCR!\Run LinXi Macro.mcr" >nul
if errorlevel 1 (
    echo [ 错误 ] 复制 LinXi.bas 失败。
    pause
    exit /b 1
)
echo [ OK ] LinXi.bas 已部署到:
echo      !DST_MCR!\Run LinXi Macro.mcr
echo.

copy /y "%SRC_VB2%" "!DST_MCR!\Define LinXi Macro.mcr" >nul
if errorlevel 1 (
    echo [ 错误 ] 复制 Me.bas 失败。
    pause
    exit /b 1
)
echo [ OK ] Me.bas 已部署到:
echo      !DST_MCR!\Define LinXi Macro.mcr
echo.

copy /y "%SRC_VB3%" "!DST_MCR!\Check LinXi Version.mcr" >nul
if errorlevel 1 (
    echo [ 错误 ] 复制 She.bas 失败。
    pause
    exit /b 1
)
echo [ OK ] She.bas 已部署到:
echo      !DST_MCR!\Check LinXi Version.mcr
echo.

rem ---- 逐个 CST 版本部署核心库与监视器宏 ----
set /a k=1

:install_loop
if !k! gtr !idx! goto :install_done

set "CSTDIR="
for %%i in (!k!) do set "CSTDIR=!VER%%i!"
if defined CSTDIR if "!CSTDIR:~-1!"=="\" set "CSTDIR=!CSTDIR:~0,-1!"

echo ============================================================
echo ---------- [ !k!/!idx! ] CST 安装根目录: !CSTDIR! ----------
echo.

set "DST_DLL=!CSTDIR!\AMD64"

echo 目标部署目录:
echo     [ LinXi.dll       ]  !DST_DLL!
echo     [ LinXi.ini       ]  !DST_DLL!
echo     [ LinXi_Watch.bas ]  !DST_DLL!
echo.

if not exist "!DST_DLL!" mkdir "!DST_DLL!"

copy /y "%SRC_LIB%" "!DST_DLL!\LinXi.dll" >nul
if errorlevel 1 (
    echo [ 错误 ] 复制 LinXi.dll 失败。
    pause
    exit /b 1
)
echo [ OK ] LinXi.dll 已部署到:
echo      !DST_DLL!\LinXi.dll
echo.

copy /y "%SRC_INI%" "!DST_DLL!\LinXi.ini" >nul
if errorlevel 1 (
    echo [ 错误 ] 复制 LinXi.ini 失败。
    pause
    exit /b 1
)
echo [ OK ] LinXi.ini 已部署到:
echo      !DST_DLL!\LinXi.ini
echo.

copy /y "%SRC_VB1%" "!DST_DLL!\LinXi_Watch.bas" >nul
if errorlevel 1 (
    echo [ 错误 ] 复制 LinXi_Watch.bas 失败。
    pause
    exit /b 1
)
echo [ OK ] LinXi.bas 已部署为监视器宏:
echo      !DST_DLL!\LinXi_Watch.bas
echo.

set /a k+=1
goto :install_loop

:install_done
echo ============================================================
echo [ 完成 ] 部署成功！共处理 !idx! 个 CST 版本, 无需任何手动选择。
echo           如需卸载, 运行同目录下的 "双击我自动卸载.bat"。
echo ============================================================
echo.
pause
exit /b 0

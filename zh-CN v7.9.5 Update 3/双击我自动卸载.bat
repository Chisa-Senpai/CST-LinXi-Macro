@echo off
setlocal EnableDelayedExpansion
chcp 936 >nul
cd /d "%~dp0"
title CST 慢波结构用户监视器 MacroKit 一键卸载

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ 提示 ] 需要管理员权限，正在重新启动...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo ==============================================
echo     CST 慢波结构用户监视器 MacroKit 一键卸载
echo ==============================================
echo.

set "TMPFILE=%~dp0_cst_detect.tmp"
powershell -NoProfile -Command "Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*CST Studio Suite*' -and $_.DisplayName -match '202[4-9]' } | Sort-Object DisplayName -Descending | ForEach-Object { if ($_.InstallLocation) { $_.InstallLocation.TrimEnd('\') } elseif ($_.DisplayIcon) { (Split-Path $_.DisplayIcon).TrimEnd('\') } } | Select-Object -Unique | Out-File -FilePath \"%~dp0_cst_detect.tmp\" -Encoding ascii"

set "CSTDIR="
set "idx=0"
set "PREV="
echo 正在检测 CST 安装路径...
echo.
for /f "usebackq delims=" %%p in ("%TMPFILE%") do (
    if not "%%p"=="!PREV!" (
        set /a idx+=1
        set "VER!idx!=%%p"
        echo   [ !idx! ] %%p
        set "PREV=%%p"
    )
)
echo.
del "%TMPFILE%" >nul 2>&1

if !idx!==0 (
    echo [ 提示 ] 未检测到 CST Studio Suite 2024 及以上版本。
    echo         仍可继续清理用户宏目录, 下面将只处理 %APPDATA% 下的内容。
    echo.
    goto :clean_macros
)

echo 检测到 !idx! 个 CST 版本，将全部卸载。
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
    echo   [ OK ] 已删除 LinXi.dll
) else (
    echo   [ 提示 ] 未发现 LinXi.dll，跳过删除。
)

if exist "!DST_DLL!\LinXi.ini" (
    del /q "!DST_DLL!\LinXi.ini"
    echo   [ OK ] 已删除 LinXi.ini
) else (
    echo   [ 提示 ] 未发现 LinXi.ini，跳过删除。
)

if exist "!DST_DLL!\LinXi_Watch.bas" (
    del /q "!DST_DLL!\LinXi_Watch.bas"
    echo   [ OK ] 已删除 LinXi_Watch.bas
) else (
    echo   [ 提示 ] 未发现 LinXi_Watch.bas，跳过删除。
)

echo.
set /a k+=1
goto :uninstall_loop

:clean_macros
set "DST_MACROS=%APPDATA%\Dassault Systemes\CST STUDIO SUITE\Library\Macros"
set "DST_MCR=!DST_MACROS!\Admin LinXi Macro"

echo ---------- 用户宏目录 ----------
if exist "!DST_MCR!" (
    rd /s /q "!DST_MCR!"
    if exist "!DST_MCR!" (
        echo   [ 错误 ] 删除失败: !DST_MCR!
        echo            目录可能被 CST 占用, 请关闭 CST 后重试。
    ) else (
        echo   [ OK ] 已删除整个宏目录: !DST_MCR!
    )
) else (
    echo   [ 提示 ] 未发现 !DST_MCR!，跳过删除。
)
echo.

echo [ 完成 ] 卸载成功！
echo.
pause

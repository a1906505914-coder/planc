@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set "PY="
if not defined PY (
    if exist "%USERPROFILE%\WorkBuddy\2026-07-15-15-26-19\.ocr_dml_venv\Scripts\python.exe" (
        set "PY=%USERPROFILE%\WorkBuddy\2026-07-15-15-26-19\.ocr_dml_venv\Scripts\python.exe"
    )
)
for /f "delims=" %%d in ('dir /b /ad /o-n "%USERPROFILE%\.workbuddy\binaries\python\versions" 2^>nul') do (
    if not defined PY (
        if exist "%USERPROFILE%\.workbuddy\binaries\python\versions\%%d\python.exe" (
            set "PY=%USERPROFILE%\.workbuddy\binaries\python\versions\%%d\python.exe"
        )
    )
)
if not defined PY (
    echo ERROR: Python not found.
    pause
    exit /b 1
)

if "%~1"=="" (
    echo 用法: 脱敏.bat ^<类型^> [文件夹]
    echo   类型: wp(账套底稿批量) / confirm(回函) / contract(借款合同) / fa(固定资产台账) / financing(融资台账) / check(入口检查)
    echo   缺省文件夹 = 拖入 或 各类型默认目录
    echo.
    pause
    exit /b 0
)

set "TYPE=%~1"
set "ARG=%~2"

if "%TYPE%"=="wp" (
    "%PY%" -B mask_dict.py --full
    "%PY%" -B mask_batch.py
) else if "%TYPE%"=="confirm" (
    if "%ARG%"=="" ( set "ARG=D:/底稿测试/AZ/数据/2025/回函" )
    "%PY%" -B confirm_mask_local.py "%ARG%"
) else if "%TYPE%"=="contract" (
    if "%ARG%"=="" ( set "ARG=D:/底稿测试/ADF/数据/2026/借款合同" )
    "%PY%" -B contract_mask_local.py "%ARG%"
) else if "%TYPE%"=="fa" (
    if "%ARG%"=="" ( set "ARG=D:/底稿测试" )
    "%PY%" -B fa_ledger_mask.py "%ARG%"
) else if "%TYPE%"=="financing" (
    if "%ARG%"=="" ( set "ARG=D:/底稿测试" )
    "%PY%" -B financing_ledger.py "%ARG%"
) else if "%TYPE%"=="check" (
    if "%ARG%"=="" ( set "ARG=D:/底稿测试" )
    "%PY%" -B mask_gate.py --dir "%ARG%"
) else (
    echo 未知类型: %TYPE%
    pause
    exit /b 1
)
echo.
echo Done.
pause

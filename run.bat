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
    echo 用法: 拖入数据文件夹 或 命令行: run.bat ^<模块^> ^<文件夹^>
    echo   模块: regen(全量) / ca / tax / payroll / bank / equity / inventory /
    echo         longterm / loan / rd / revenue / expense / pl / gp
    echo   （模块缺省 = regen 全量重跑）
    echo.
    pause
    exit /b 0
)

set "MOD=%~1"
set "ARG=%~2"
if "%ARG%"=="" (
    set "ARG=%~1"
    set "MOD=regen"
)

if "%MOD%"=="regen" (
    "%PY%" -B regen_all.py "%ARG%"
) else (
    "%PY%" -B launcher.py %MOD% "%ARG%"
)
echo.
echo Done.
pause

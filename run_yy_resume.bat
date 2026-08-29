@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

REM Locate Python: prefer WorkBuddy managed python, then py launcher.
REM NOTE: never use 'goto' inside a for loop - it corrupts block parsing and
REM silently crashes the batch. Use a top-level for with 'if not defined PY'
REM to take the first matching version. Keep this file ASCII-only so it parses
REM correctly under any codepage (no BOM dependency).
REM Prefer project venv (full deps; managed 3.13 openpyxl broken).
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
    where py >nul 2>nul && set "PY=py"
)
if not defined PY (
    echo ============================================================
    echo ERROR: Python not found.
    echo This tool uses the WorkBuddy bundled Python.
    echo Please make sure WorkBuddy is installed correctly.
    echo ============================================================
    pause
    exit /b 1
)

echo ============================================
echo   yy 账套断点续跑 - 双击运行
echo   日志: <DATA_ROOT>/yy/_yy_resume.log
echo   跑完前请勿关闭本窗口
echo ============================================
"%PY%" -B resume_yy.py
echo.
echo 完成，按任意键关闭
pause >nul

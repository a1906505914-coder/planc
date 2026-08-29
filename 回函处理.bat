@echo off
chcp 65001 >nul
title 回函全流程处理（脱敏第一步）
echo ============================================
echo   回函全流程处理
echo   [第1步] 本地先行脱敏（真实名仅本机处理，不对外）
echo   [第2步] 解析生成 脱敏明细表.xlsx
echo ============================================
echo.
set "VENV_PY=C:\Users\lvlh\WorkBuddy\2026-07-15-15-26-19\.ocr_dml_venv\Scripts\python.exe"
set "APP=D:\底稿测试\账套取数审计小程序"

echo [1/2] 本地先行脱敏...
"%VENV_PY%" "%APP%\confirm_mask_local.py"
if errorlevel 1 goto :err
echo.
echo [2/2] 解析生成明细表...
"%VENV_PY%" "%APP%\confirm_parse_all.py" AZ --year 2025
if errorlevel 1 goto :err
echo.
echo ============================================
echo   完成。可外发成果:
echo     D:\底稿测试\AZ\数据\2025\回函\回函_脱敏明细表.xlsx
echo     D:\底稿测试\AZ\数据\2025\回函\回函_解析结果_脱敏.json
echo   真实名对照表仅存本机: confirm_mask.json
echo ============================================
pause
exit /b 0

:err
echo [错误] 处理失败，请查看上方提示。
pause
exit /b 1

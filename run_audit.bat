@echo off
REM 审计小程序统一启动器（固定 PYTHONHASHSEED → 生成器行序稳定，消除"重跑行序不同"的虚假回退）
REM 用法：run_audit.bat <python脚本> [参数...]
REM 例：  run_audit.bat run_u8_on_sap.py "D:/数据包" --group
REM 例：  run_audit.bat current_account_detail.py "D:/数据包"
setlocal
set PYTHONHASHSEED=0
set AH_PY=%~dp0\..\.workbuddy\binaries\python\versions\3.13.12\python.exe
if exist "%AH_PY%" (
    "%AH_PY%" -X utf8 %*
) else (
    python -X utf8 %*
)
endlocal

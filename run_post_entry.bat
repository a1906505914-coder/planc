@echo off
chcp 65001 >nul
echo 审计调整分录过入器（post_entry）
echo 用法1：把《审计调整分录汇总表.xlsx》拖到本窗口后回车 → 过入到该文件所在文件夹的底稿
echo 用法2：拖入分录表后，再输入底稿文件夹路径后回车
set /p ENTRY=分录表文件：
set /p FOLDER=底稿文件夹（直接回车=分录表所在目录）：
if "%FOLDER%"=="" (
  python -B "%~dp0post_entry.py" "%ENTRY%"
) else (
  python -B "%~dp0post_entry.py" "%ENTRY%" "%FOLDER%"
)
pause

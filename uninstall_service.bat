@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   考勤报表自动导出 - 卸载后台服务
echo ============================================
echo.

set "TASK_NAME=KaoqinAutoReport"

echo 正在停止服务...
python auto_report\service_wrapper.py --stop >nul 2>nul

echo 正在删除计划任务...
schtasks /delete /tn "%TASK_NAME%" /f >nul 2>nul

echo.
echo [完成] 服务已卸载
echo.
pause

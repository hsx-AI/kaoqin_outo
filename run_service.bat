@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 考勤报表导出服务 - 保持开启

:loop
echo [%date% %time%] 正在启动服务...
python auto_report/http_service.py --host 0.0.0.0 --port 6648
echo [%date% %time%] 服务异常退出，5秒后尝试重启...
timeout /t 5 >nul
goto loop

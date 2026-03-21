@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   考勤报表自动导出 - 后台服务安装
echo ============================================
echo.
echo 注意: 请以管理员身份运行此脚本!
echo.

set "TASK_NAME=KaoqinAutoReport"
set "WRAPPER_SCRIPT=%~dp0auto_report\service_wrapper.py"

:: 检查 service_wrapper.py 是否存在
if not exist "%WRAPPER_SCRIPT%" (
    echo [错误] 未找到 %WRAPPER_SCRIPT%
    pause
    exit /b 1
)

:: 查找 pythonw.exe 路径
set "PYTHONW_PATH="
where pythonw.exe >nul 2>nul
if %errorlevel% equ 0 (
    for /f "delims=" %%i in ('where pythonw.exe') do (
        if not defined PYTHONW_PATH set "PYTHONW_PATH=%%i"
    )
)

if not defined PYTHONW_PATH (
    echo [警告] 未在PATH中找到 pythonw.exe
    echo 尝试从 python.exe 推断路径...
    for /f "delims=" %%i in ('where python.exe 2^>nul') do (
        set "PYTHON_DIR=%%~dpi"
    )
    if defined PYTHON_DIR (
        if exist "%PYTHON_DIR%pythonw.exe" (
            set "PYTHONW_PATH=%PYTHON_DIR%pythonw.exe"
        )
    )
)

if not defined PYTHONW_PATH (
    echo [错误] 无法找到 pythonw.exe, 请确保 Python 已安装并在 PATH 中
    pause
    exit /b 1
)

echo 使用 Python: %PYTHONW_PATH%
echo 服务脚本:    %WRAPPER_SCRIPT%
echo.

:: 停止已有服务
echo 正在停止已有服务...
python "%WRAPPER_SCRIPT%" --stop >nul 2>nul

:: 删除旧的计划任务
schtasks /delete /tn "%TASK_NAME%" /f >nul 2>nul

:: 使用 PowerShell 创建计划任务（支持崩溃重启等高级选项）
echo 正在创建计划任务...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$action = New-ScheduledTaskAction -Execute '\"%PYTHONW_PATH%\"' -Argument '\"%WRAPPER_SCRIPT%\"' -WorkingDirectory '%~dp0auto_report'; ^
     $trigger = New-ScheduledTaskTrigger -AtLogOn; ^
     $settings = New-ScheduledTaskSettingsSet -RestartCount 9999 -RestartInterval (New-TimeSpan -Seconds 30) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Days 3650) -StartWhenAvailable; ^
     $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest -LogonType Interactive; ^
     Register-ScheduledTask -TaskName '%TASK_NAME%' -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force"

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   [成功] 服务安装完成!
    echo ============================================
    echo.
    echo   任务名称:  %TASK_NAME%
    echo   启动方式:  用户登录后自动启动
    echo   崩溃恢复:  30秒后自动重启
    echo   运行模式:  后台无窗口
    echo.
    echo   管理命令:
    echo     查看状态:  python auto_report\service_wrapper.py --status
    echo     停止服务:  python auto_report\service_wrapper.py --stop
    echo     启动服务:  通过任务计划程序启动, 或直接运行:
    echo                start "" "%PYTHONW_PATH%" "%WRAPPER_SCRIPT%"
    echo.
    echo   正在立即启动服务...
    start "" "%PYTHONW_PATH%" "%WRAPPER_SCRIPT%"
    timeout /t 2 >nul
    python "%WRAPPER_SCRIPT%" --status
) else (
    echo.
    echo [失败] 计划任务创建失败, 请确认以管理员身份运行
)

echo.
pause

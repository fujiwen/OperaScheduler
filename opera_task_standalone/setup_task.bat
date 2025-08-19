@echo off
chcp 65001 >nul
echo ========================================
echo 设置Opera DataGuard计划任务
echo ========================================

echo.
echo 此脚本将帮助您设置Windows计划任务
echo 默认设置为每天上午8:00执行
echo.

set /p TASK_PATH="请输入OperaDataGuardTask.exe的完整路径: "
if not exist "%TASK_PATH%" (
    echo 错误: 文件不存在 - %TASK_PATH%
    pause
    exit /b 1
)

set TASK_DIR=%~dp1
if "%TASK_DIR%"=="" (
    for %%i in ("%TASK_PATH%") do set TASK_DIR=%%~dpi
)

echo.
echo 正在创建计划任务...
echo 任务名称: OperaDataGuardTask
echo 执行文件: %TASK_PATH%
echo 工作目录: %TASK_DIR%
echo 执行时间: 每天上午8:00
echo.

:: 删除已存在的任务（如果有）
schtasks /delete /tn "OperaDataGuardTask" /f >nul 2>&1

:: 创建新的计划任务
schtasks /create /tn "OperaDataGuardTask" ^^
    /tr "\"%TASK_PATH%\"" ^^
    /sc daily ^^
    /st 08:00 ^^
    /sd %date% ^^
    /ru "SYSTEM" ^^
    /rl highest ^^
    /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo 计划任务创建成功！
    echo.
    echo 任务管理命令:
    echo   查看任务: schtasks /query /tn "OperaDataGuardTask"
    echo   运行任务: schtasks /run /tn "OperaDataGuardTask"
    echo   停止任务: schtasks /end /tn "OperaDataGuardTask"
    echo   删除任务: schtasks /delete /tn "OperaDataGuardTask" /f
    echo.
    echo 您也可以通过"任务计划程序"图形界面管理此任务
) else (
    echo.
    echo 计划任务创建失败！
    echo 请检查权限或手动创建计划任务
)

echo.
pause
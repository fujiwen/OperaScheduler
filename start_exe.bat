@echo off
chcp 65001
set PYTHONIOENCODING=utf-8
echo 正在启动Opera数据库监控工具...

REM 检查是否存在目录模式的可执行文件
if exist "dist\Opera数据库监控工具\Opera数据库监控工具.exe" (
    start "" "dist\Opera数据库监控工具\Opera数据库监控工具.exe"
) else if exist "dist\Opera数据库监控工具_单文件版.exe" (
    start "" "dist\Opera数据库监控工具_单文件版.exe"
) else (
    echo 未找到可执行文件，请先运行build_exe.bat构建程序。
    pause
    exit /b 1
)

echo 程序已启动。
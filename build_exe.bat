@echo off
chcp 65001
set PYTHONIOENCODING=utf-8
echo ===================================
echo  Opera数据库监控工具 - 构建脚本
echo ===================================
echo.

REM 检查是否已安装PyInstaller
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller未安装，正在安装...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo PyInstaller安装失败，请手动安装后重试。
        pause
        exit /b 1
    )
    echo PyInstaller安装成功。
)

echo.
echo 开始构建Opera数据库监控工具...
echo.

REM 创建构建目录
if not exist build mkdir build
if not exist dist mkdir dist

REM 执行PyInstaller构建命令
pyinstaller --clean ^^
    --name="Opera数据库监控工具" ^^
    --windowed ^^
    --add-data="check_standby.bat;." ^^
    --add-data="check_standby.sql;." ^^
    --add-data="daily_report.bat;." ^^
    --add-data="daily_report_dg.sql;." ^^
    --add-data="daily_report_prod.sql;." ^^
    --add-data="start_standby.sql;." ^^
    --add-data="opera_monitor.ini;." ^^
    --add-data="logs;logs" ^^
    opera_monitor.py

if %errorlevel% neq 0 (
    echo 构建失败，请检查错误信息。
    pause
    exit /b 1
)

echo.
echo 构建完成！可执行文件位于 dist\Opera数据库监控工具 目录中。
echo.

REM 复制logs目录到dist目录
if not exist "dist\Opera数据库监控工具\logs" mkdir "dist\Opera数据库监控工具\logs"

echo 是否要构建单文件版本？(Y/N)
set /p choice="选择: "
if /i "%choice%"=="Y" (
    echo.
    echo 开始构建单文件版本...
    echo.
    
    pyinstaller --clean ^^
        --name="Opera数据库监控工具_单文件版" ^^
        --windowed ^^
        --onefile ^^
        --add-data="check_standby.bat;." ^^
        --add-data="check_standby.sql;." ^^
        --add-data="daily_report.bat;." ^^
        --add-data="daily_report_dg.sql;." ^^
        --add-data="daily_report_prod.sql;." ^^
        --add-data="start_standby.sql;." ^^
        --add-data="opera_monitor.ini;." ^^
        --add-data="logs;logs" ^^
        opera_monitor.py
    
    if %errorlevel% neq 0 (
        echo 单文件版本构建失败，请检查错误信息。
    ) else (
        echo.
        echo 单文件版本构建完成！可执行文件位于 dist 目录中。
        echo.
    )
)

echo 构建过程已完成。
pause
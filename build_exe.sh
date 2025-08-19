#!/bin/bash

echo "==================================="
echo " Opera数据库监控工具 - 构建脚本"
echo "==================================="
echo 

# 检查是否已安装PyInstaller
if ! pip show pyinstaller > /dev/null 2>&1; then
    echo "PyInstaller未安装，正在安装..."
    pip install pyinstaller
    if [ $? -ne 0 ]; then
        echo "PyInstaller安装失败，请手动安装后重试。"
        read -p "按回车键继续..."
        exit 1
    fi
    echo "PyInstaller安装成功。"
fi

echo 
echo "开始构建Opera数据库监控工具..."
echo 

# 创建构建目录
mkdir -p build dist

# 执行PyInstaller构建命令
pyinstaller --clean \
    --name="Opera数据库监控工具" \
    --windowed \
    --add-data="check_standby.bat:." \
    --add-data="check_standby.sql:." \
    --add-data="daily_report.bat:." \
    --add-data="daily_report_dg.sql:." \
    --add-data="daily_report_prod.sql:." \
    --add-data="start_standby.sql:." \
    --add-data="opera_monitor.ini:." \
    --add-data="logs:logs" \
    opera_monitor.py

if [ $? -ne 0 ]; then
    echo "构建失败，请检查错误信息。"
    read -p "按回车键继续..."
    exit 1
fi

echo 
echo "构建完成！可执行文件位于 dist/Opera数据库监控工具 目录中。"
echo 

# 复制logs目录到dist目录
mkdir -p "dist/Opera数据库监控工具/logs"

echo "是否要构建单文件版本？(Y/N)"
read -p "选择: " choice
if [[ "$choice" == "Y" || "$choice" == "y" ]]; then
    echo 
    echo "开始构建单文件版本..."
    echo 
    
    pyinstaller --clean \
        --name="Opera数据库监控工具_单文件版" \
        --windowed \
        --onefile \
        --add-data="check_standby.bat:." \
        --add-data="check_standby.sql:." \
        --add-data="daily_report.bat:." \
        --add-data="daily_report_dg.sql:." \
        --add-data="daily_report_prod.sql:." \
        --add-data="start_standby.sql:." \
        --add-data="opera_monitor.ini:." \
        --add-data="logs:logs" \
        opera_monitor.py
    
    if [ $? -ne 0 ]; then
        echo "单文件版本构建失败，请检查错误信息。"
    else
        echo 
        echo "单文件版本构建完成！可执行文件位于 dist 目录中。"
        echo 
    fi
fi

echo "构建过程已完成。"
read -p "按回车键继续..."
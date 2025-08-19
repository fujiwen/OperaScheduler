# Opera数据库监控工具构建指南

本文档提供了如何使用PyInstaller将Opera数据库监控工具打包成可执行文件的详细说明。

## 前提条件

- 已安装Python 3.6或更高版本
- 已安装pip包管理器
- 已安装项目依赖（可通过`pip install -r requirements.txt`安装）

## 构建脚本

本项目提供了两个构建脚本：

1. `build_exe.bat` - 用于Windows平台
2. `build_exe.sh` - 用于macOS/Linux平台

## Windows平台构建步骤

1. 打开命令提示符或PowerShell
2. 导航到项目目录：`cd 路径\到\OperaScheduler`
3. 运行构建脚本：`build_exe.bat`
4. 脚本会自动检查并安装PyInstaller（如果尚未安装）
5. 构建完成后，可执行文件将位于`dist\Opera数据库监控工具`目录中
6. 脚本会询问是否构建单文件版本，根据需要选择Y或N
7. 使用`start_exe.bat`启动打包后的程序，以避免乱码问题

### Windows平台乱码和错误问题解决

在Windows平台上，可能会遇到乱码和错误问题。我们已经在以下几个方面进行了优化：

1. 所有文件读写操作都使用UTF-8编码
2. 配置文件读写使用UTF-8编码
3. 日志文件使用UTF-8编码
4. 批处理文件执行时使用UTF-8编码
5. 启动脚本设置了控制台代码页为UTF-8(65001)

如果仍然遇到问题，请参考`WINDOWS_README.md`文件获取更详细的解决方法。

## macOS/Linux平台构建步骤

1. 打开终端
2. 导航到项目目录：`cd 路径/到/OperaScheduler`
3. 给构建脚本添加执行权限：`chmod +x build_exe.sh`
4. 运行构建脚本：`./build_exe.sh`
5. 脚本会自动检查并安装PyInstaller（如果尚未安装）
6. 构建完成后，可执行文件将位于`dist/Opera数据库监控工具`目录中
7. 脚本会询问是否构建单文件版本，根据需要选择Y或N

## 构建选项说明

### 目录模式（默认）

- 生成一个包含主可执行文件和所有依赖的目录
- 启动速度快
- 需要保持目录结构完整

### 单文件模式

- 生成单个可执行文件
- 便于分发
- 启动速度较慢（需要解压缩到临时目录）

## 常见问题及解决方案

### 1. tkinter相关错误

**问题**：构建时或运行时出现tkinter相关错误

**解决方案**：
- Windows：确保安装Python时选择了包含tkinter的选项
- macOS：`brew install python-tk@3.9`（根据Python版本调整）
- Linux：`sudo apt-get install python3-tk`（Ubuntu/Debian）或`sudo dnf install python3-tkinter`（Fedora）

### 2. 找不到依赖文件

**问题**：运行打包后的程序时提示找不到某些文件

**解决方案**：
- 确保构建命令中包含了所有必要的`--add-data`参数
- 检查程序中的文件路径处理逻辑，确保兼容PyInstaller打包环境
- 特别注意logs目录已被包含在构建中，程序可以正常访问该目录

### 3. 运行时出现"Failed to execute script"错误

**问题**：程序无法启动，提示脚本执行失败

**解决方案**：
- 尝试在命令行中运行可执行文件，查看详细错误信息
- 添加`--debug=all`参数重新构建，以获取更多调试信息

## 自定义构建

如果需要自定义构建过程，可以直接编辑构建脚本或使用以下命令手动构建：

```
pyinstaller --name="Opera数据库监控工具" --windowed [其他选项] opera_monitor.py
```

可用的其他选项包括：

- `--icon=图标文件.ico` - 设置应用程序图标
- `--version-file=版本文件.txt` - 添加版本信息
- `--noconsole` - 不显示控制台窗口（与--windowed相同）
- `--add-data="源文件;目标目录"` - 添加数据文件（Windows使用分号分隔，macOS/Linux使用冒号）

## 注意事项

1. 构建的可执行文件仅适用于构建它的操作系统
2. 如果程序需要访问外部文件，确保这些文件在运行时可用
3. 单文件模式下，临时文件会解压到系统临时目录，可能影响性能
4. 确保所有批处理文件（.bat）在Windows环境中有正确的行结束符（CRLF）
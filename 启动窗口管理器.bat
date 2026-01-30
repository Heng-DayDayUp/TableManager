@echo off

:: 窗口管理器启动脚本
:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python。请先安装Python 3.6或更高版本。
    pause
    exit /b 1
)

:: 切换到脚本目录
cd /d "%~dp0"

:: 启动应用
echo 正在启动窗口管理器...
python app.py

:: 等待用户按任意键退出
pause

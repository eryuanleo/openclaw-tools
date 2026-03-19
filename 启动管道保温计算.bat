@echo off
chcp 65001 >nul
title 化工管道保温表面积计算程序

echo ================================================
echo       化工管道保温表面积计算程序
echo ================================================
echo.

REM 检查Python是否安装
echo 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到Python！
    echo.
    echo 请先安装Python 3.x：
    echo 1. 访问 https://www.python.org/downloads/
    echo 2. 下载并安装Python
    echo 3. 安装时务必勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

REM 检查openpyxl库
echo 检查所需库...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ⚠ 缺少openpyxl库，正在安装...
    pip install openpyxl
    if errorlevel 1 (
        echo ❌ 安装失败，请手动安装：
        echo pip install openpyxl
        pause
        exit /b 1
    )
    echo ✓ openpyxl安装成功
)

echo.
echo ✓ 环境检查通过
echo.

REM 运行主程序
echo 正在启动管道保温计算程序...
echo.
python pipe_insulation_calculator.py

echo.
echo 程序执行完毕！
echo.
pause
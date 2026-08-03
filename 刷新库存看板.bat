@echo off
chcp 65001 >nul
title ITO 库存看板 · 每日刷新
setlocal

cd /d "E:\ITO库存看板系统"

echo ================================================
echo   ITO 库存看板 · 每日数据刷新
echo   数据源: 库存-每日.xlsx
echo   输出 : ITO库存看板.html（双击即可打开看板）
echo ================================================
echo.

if not exist "库存-每日.xlsx" (
    echo [错误] 找不到「库存-每日.xlsx」
    echo 请确认每日更新的库存表已放到: %cd%
    echo.
    pause
    exit /b 1
)

echo [1/3] 正在检测 Python 环境...
set PYCMD=
py -3.11 -c "import openpyxl" >nul 2>&1 && set PYCMD=py -3.11
if not defined PYCMD py -c "import openpyxl" >nul 2>&1 && set PYCMD=py
if not defined PYCMD python -c "import openpyxl" >nul 2>&1 && set PYCMD=python
if not defined PYCMD (
    echo [错误] 未找到可用的 Python 环境（需要 openpyxl 模块）
    echo 请安装:  pip install openpyxl
    echo.
    pause
    exit /b 1
)
echo       使用: %PYCMD%

echo [2/3] 正在读取库存数据并计算指标...
%PYCMD% build_inventory.py
if errorlevel 1 goto :ERR

echo.
echo [3/3] 刷新完成！
echo        看板成品: %cd%\ITO库存看板.html
echo        正在打开看板...
start "" "ITO库存看板.html"

echo.
echo 完成。下次刷新：直接双击本脚本，或用任务计划程序每日定时运行。
echo.
pause
exit /b 0

:ERR
echo.
echo [错误] 刷新失败！请检查：
echo   1. 库存-每日.xlsx 是否在本目录且未被占用（Excel 打开中请先关闭）
echo   2. 文件格式是否为 ERP 导出的库存表（含商家编码表头）
echo.
pause
exit /b 1

@echo off
title ITO 库存看板 · 每日刷新
setlocal

cd /d "D:\E盘文件\ITO库存看板系统"

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

echo [1/4] 正在检测 Python 环境...
set PYCMD=C:\Users\QwQ\.workbuddy\binaries\python\versions\3.13.12\python.exe
"%PYCMD%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到可用的 Python 环境（需要 openpyxl 模块）
    echo 请安装:  pip install openpyxl
    echo.
    pause
    exit /b 1
)
echo       使用: %PYCMD%

set PYTHONIOENCODING=gbk
echo [2/4] 正在读取库存数据并计算指标...
%PYCMD% build_inventory.py
if errorlevel 1 goto :ERR
echo [%date% %time%] build OK >> run.log

echo.
echo [3/4] 正在推送钉钉预警...
%PYCMD% send_dingtalk.py
if errorlevel 1 echo [提示] 钉钉推送失败（不影响本地构建）
echo [%date% %time%] dingtalk done >> run.log

echo.
echo [3.5/4] 更新时间轴已由 build 流程自动注入（含 deploy/index.html 同步）

echo.
echo [4/4] 正在更新网页版（Vercel）...
if not exist "vercel_token.txt" (
    echo [提示] 未找到 vercel_token.txt，跳过网页部署（本地看板已更新）
    goto :SKIPDEPLOY
)
echo        正在部署...
REM ===== 修复 2026-08-19：改用独立 Python 部署脚本（精确退出码 + 失败自动重试 3 次 + 完整日志），根治 bat errorlevel 误判 =====
set "NODE_OPTIONS="
%PYCMD% "%~dp0deploy_vercel.py"
if errorlevel 1 (
    echo [提示] 网页部署 3 次尝试均失败，详见 deploy_vercel_fail_*.log，本地看板已更新
    echo [%date% %time%] deploy FAIL >> run.log
) else (
    echo [ok] 网页版已更新: https://ito-inventory-dashboard.vercel.app
    echo [%date% %time%] deploy OK >> run.log
)
:SKIPDEPLOY

echo.
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

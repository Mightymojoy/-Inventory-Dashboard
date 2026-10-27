@echo off
title ITO 库存看板 · 每日刷新
setlocal

REM █ 修复：用 %~dp0 替代项目绝对路径，项目搬家/改名不再失效 █
cd /d "%~dp0"

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
REM ====================================================================
REM 动态定位 Python.exe【禁止写死版本号或安装目录】
REM   平台重装会删除 envs\default；versions\ 下是符号链接且版本号会变。
REM   优先级：环境变量 ITO_PYTHON -> venv(唯一带 pandas/openpyxl) -> versions\current
REM   绝不兜底 WindowsApps\python.exe（商店 stub，运行会弹窗卡住无人值守）
REM ====================================================================
set "PYTHON="
set "PYVENV=C:\Users\QwQ\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
set "PYVER=C:\Users\QwQ\.workbuddy\binaries\python\versions"
set "PYCV="

if defined ITO_PYTHON if exist "%ITO_PYTHON%" set "PYTHON=%ITO_PYTHON%"
if defined PYTHON goto PY_READY

if exist "%PYVENV%" set "PYTHON=%PYVENV%"
if defined PYTHON goto PY_READY

if exist "%PYVER%\current" for /f "usebackq delims=" %%V in ("%PYVER%\current") do set "PYCV=%%V"
if defined PYCV if exist "%PYVER%\%PYCV%\python.exe" set "PYTHON=%PYVER%\%PYCV%\python.exe"
if defined PYTHON goto PY_READY

echo [X] 找不到可用的 Python.exe！
echo     已尝试：环境变量 ITO_PYTHON、venv、versions\current
echo     修复建议：重装 WorkBuddy 后重建虚拟环境
echo     本次未执行任何生成/推送，数据未被改动。
pause
exit /b 1

:PY_READY
echo [i] 使用 Python: %PYTHON%
if /i not "%PYTHON%"=="%PYVENV%" (
    echo [!] 警告：当前不是标准 venv，可能缺少 pandas/openpyxl
    echo     若报 ModuleNotFoundError，请重建 venv 后重跑本脚本
)
echo.
"%PYTHON%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [X] 当前 Python 缺少 openpyxl 模块，无法读取 Excel。
    echo     请运行: pip install openpyxl
    echo     本次未执行任何生成/部署，数据未被改动。
    echo.
    pause
    exit /b 1
)
echo       依赖检查通过（openpyxl 可用）

set PYTHONIOENCODING=gbk
echo [2/4] 正在读取库存数据并计算指标...
"%PYTHON%" build_inventory.py
if errorlevel 1 goto :ERR
echo [%date% %time%] build OK >> run.log

echo.
echo [3/4] 正在推送钉钉预警...
"%PYTHON%" send_dingtalk.py
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
"%PYTHON%" "%~dp0deploy_vercel.py"
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

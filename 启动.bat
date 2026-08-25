@echo off
setlocal
title 洛阳师范学院综测计算程序
cd /d "%~dp0app"

REM 自动检测 Python（兼容不同设备）
set "PY="
for /f "delims=" %%i in ('where py 2^>nul') do set "PY=py"
if not defined PY for /f "delims=" %%i in ('where python 2^>nul') do set "PY=python"
if not defined PY (
    echo 未找到 Python，请先安装 Python 3 并加入 PATH。
    pause
    exit /b 1
)

set "URL=http://127.0.0.1:5174"

echo ============================================
echo   洛阳师范学院综测计算程序
echo   2024级软件工程十班 李祎杭
echo   v2.0
echo ============================================
echo.
echo  [1/3] 正在启动服务器...
start "ZongceCalc-Server" /min %PY% "server.py"

echo  [2/3] 正在等待服务器就绪...
ping -n 5 127.0.0.1 >nul

echo  [3/3] 正在打开浏览器 %URL% ...
start "" "%URL%"

echo.
echo  完成！如果浏览器提示"拒绝连接"，关闭本窗口，稍等片刻后重新双击运行即可。
echo  关闭本窗口不会退出程序（服务器仍在后台运行中）...
pause >nul

@echo off
setlocal
chcp 65001 >nul
title 打包洛阳师范学院综测计算程序
cd /d "%~dp0"

set "PY="
for /f "delims=" %%i in ('where py 2^>nul') do set "PY=py"
if not defined PY for /f "delims=" %%i in ('where python 2^>nul') do set "PY=python"
if not defined PY (
    echo 未找到 Python，请先安装 Python 3 并加入 PATH。
    pause
    exit /b 1
)

echo ============================================
echo   正在打包洛阳师范学院综测计算程序
echo ============================================
%PY% -m PyInstaller --noconfirm --clean "洛阳师范学院综测计算程序.spec"

echo.
echo 打包完成，exe 位于 dist\洛阳师范学院综测计算程序.exe
pause >nul

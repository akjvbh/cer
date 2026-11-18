@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "AUTO_OPEN_BROWSER=1"
if not defined OFFICE_TIMEOUT_SECONDS set "OFFICE_TIMEOUT_SECONDS=120"

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] 创建虚拟环境 .venv ...
    py -3.11 -m venv .venv
    if errorlevel 1 python -m venv .venv
    if errorlevel 1 (
        echo 创建虚拟环境失败：请安装 Python 3.11，并确保 py 启动器可用。
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import fastapi, uvicorn, fitz, docx, PIL, pdf2docx, reportlab, markdown" >nul 2>nul
if errorlevel 1 (
    echo [2/3] 安装依赖，首次运行可能需要几分钟 ...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo 依赖安装失败，请检查网络或 pip 源。
        pause
        exit /b 1
    )
)

set "PORT=8000"
".venv\Scripts\python.exe" -c "import socket; s=socket.socket(); s.bind(('127.0.0.1',8000)); s.close()" >nul 2>nul
if errorlevel 1 set "PORT=8001"

if "%PORT%"=="8001" echo 端口 8000 被占用，改用 8001。
echo [3/3] 启动地址：http://127.0.0.1:%PORT%
echo 关闭此窗口即可停止服务。
".venv\Scripts\python.exe" main.py
if errorlevel 1 (
    echo.
    echo 服务异常退出，请查看上方日志。
    pause
)
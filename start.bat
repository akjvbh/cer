@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

if not defined AUTO_OPEN_BROWSER set "AUTO_OPEN_BROWSER=1"
if not defined OFFICE_TIMEOUT_SECONDS set "OFFICE_TIMEOUT_SECONDS=120"
set "START_LOG=1"
set "LOG_FILE=%CD%\_start_log.txt"

> "%LOG_FILE%" echo [%date% %time%] 文档工坊启动
call :log "workdir %CD%"

if exist ".venv\Scripts\python.exe" goto :deps
echo [1/3] 创建虚拟环境 .venv ...
call :log "step1 create venv"
py -3.11 -m venv .venv
if errorlevel 1 python -m venv .venv
if exist ".venv\Scripts\python.exe" goto :deps
echo 创建虚拟环境失败，请安装 Python 3.11 并确保 py 启动器可用
call :log "step1 failed: no Python 3.11"
pause
exit /b 1

:deps
".venv\Scripts\python.exe" -c "import fastapi, uvicorn, fitz, docx, PIL, pdf2docx, reportlab, markdown" >nul 2>nul
if not errorlevel 1 goto :ports
echo [2/3] 安装依赖，首次运行可能需要几分钟 ...
call :log "step2 install dependencies"
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if not errorlevel 1 goto :ports
echo 依赖安装失败，请检查网络或 pip 源
call :log "step2 failed: pip install error"
pause
exit /b 1

:ports
set "PORT="
for %%P in (8000 8001 8002 8003 8004 8005) do if not defined PORT call :try_port %%P
if defined PORT goto :port_ready
echo 端口 8000-8005 都被占用，请先关闭占用这些端口的程序
call :log "all ports 8000-8005 busy, abort"
pause
exit /b 1

:try_port
if defined PORT exit /b 0
".venv\Scripts\python.exe" -c "import socket; s=socket.socket(); s.bind(('127.0.0.1',%1)); s.close()" >nul 2>nul
if not errorlevel 1 set "PORT=%1"
exit /b 0

:port_ready
if "%PORT%"=="8000" goto :launch
echo 端口 8000 被占用，改用 %PORT%
call :log "port 8000 busy, using %PORT%"

:launch
echo [3/3] 启动地址：http://127.0.0.1:%PORT%
echo 关闭此窗口即可停止服务
call :log "listening http://127.0.0.1:%PORT%"
".venv\Scripts\python.exe" main.py
if not errorlevel 1 exit /b 0
echo.
echo 服务异常退出，请查看上方输出或 _start_log.txt
call :log "server exited with error"
pause
exit /b 0

:log
>>"%LOG_FILE%" echo(%~1
exit /b 0

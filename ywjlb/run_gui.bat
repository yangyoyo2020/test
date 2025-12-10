@echo off
REM 运维记录簿处理工具 - Windows快速启动脚本
REM 此脚本用于快速启动图形用户界面

setlocal enabledelayedexpansion

REM 获取当前脚本所在目录
set SCRIPT_DIR=%~dp0

REM 进入脚本目录
cd /d "%SCRIPT_DIR%"

REM 查找虚拟环境
if exist "..\..\.venv\Scripts\python.exe" (
    set PYTHON_EXE=..\..\.venv\Scripts\python.exe
    echo 使用虚拟环境...
) else if exist "..\.venv\Scripts\python.exe" (
    set PYTHON_EXE=..\.venv\Scripts\python.exe
    echo 使用虚拟环境...
) else (
    set PYTHON_EXE=python
    echo 使用系统Python...
)

REM 启动应用
echo 正在启动运维记录簿处理工具...
"%PYTHON_EXE%" ywjlb_app.py %*

pause

#!/usr/bin/env pwsh
<#
.SYNOPSIS
    运维记录簿处理工具 - PowerShell启动脚本
    
.DESCRIPTION
    此脚本用于快速启动图形用户界面
    
.EXAMPLE
    .\run_gui.ps1
#>

# 获取脚本目录
$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# 进入脚本目录
Push-Location $ScriptDir

try {
    # 寻找Python可执行文件
    $PythonExe = $null
    
    # 检查虚拟环境
    if (Test-Path "..\..\..\.venv\Scripts\python.exe") {
        $PythonExe = "..\..\..\.venv\Scripts\python.exe"
        Write-Host "✓ 使用虚拟环境 (..\..\..\.venv)"
    }
    elseif (Test-Path "..\..\.venv\Scripts\python.exe") {
        $PythonExe = "..\..\.venv\Scripts\python.exe"
        Write-Host "✓ 使用虚拟环境 (..\..\.venv)"
    }
    elseif (Test-Path ".\.venv\Scripts\python.exe") {
        $PythonExe = ".\.venv\Scripts\python.exe"
        Write-Host "✓ 使用虚拟环境 (.\.venv)"
    }
    else {
        $PythonExe = "python"
        Write-Host "✓ 使用系统Python"
    }
    
    # 验证Python
    & $PythonExe --version 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "✗ 错误：找不到Python" -ForegroundColor Red
        exit 1
    }
    
    # 启动应用
    Write-Host "正在启动运维记录簿处理工具..." -ForegroundColor Green
    & $PythonExe ywjlb_app.py $args
}
catch {
    Write-Host "✗ 错误：$($_)" -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}

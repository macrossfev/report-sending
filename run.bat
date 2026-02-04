@echo off
chcp 65001 >nul
title PDF加密邮件发送系统 - Excel配置版

echo.
echo ====================================
echo  PDF加密邮件发送系统 (Excel配置版)
echo ====================================
echo.

REM 检查Python是否安装
py --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python
    pause
    exit /b 1
)

REM 运行脚本
py pdf_encrypt_send.py

pause

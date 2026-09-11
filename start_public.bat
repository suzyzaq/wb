@echo off
chcp 65001 >nul
title 纯白图识别服务 - 公域隧道（ngrok）
cd /d "%~dp0"
python launcher.py public
pause

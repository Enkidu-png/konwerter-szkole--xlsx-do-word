@echo off
chcp 65001 >nul 2>&1
title Konwerter Szkolen
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0launch.ps1"
pause

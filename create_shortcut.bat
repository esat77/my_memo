@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0create_shortcut.ps1"
echo "デスクトップにショートカットを作成しました。"
pause

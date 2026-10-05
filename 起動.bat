@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 施工計画書アプリ
echo 施工計画書アプリを起動しています...
python app.py
if errorlevel 1 (
  echo.
  echo 起動できませんでした。Python が入っているか確認してください。
  pause
)

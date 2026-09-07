@echo off
chcp 65001 >nul
title 台灣股票收盤價 Excel 自動更新工具
cd /d "%~dp0"

python update_stocks.py %*

if errorlevel 1 (
    echo.
    echo ======================================================================
    echo  執行過程中發生錯誤！詳細 Log 訊息請參閱: update_stocks.log
    echo ======================================================================
    pause
)

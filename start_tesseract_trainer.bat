@echo off
chcp 65001 >nul
title Tesseract 驗證碼訓練器

echo.
echo ====================================================
echo           🎯 Tesseract 驗證碼訓練器
echo ====================================================
echo.
echo 📋 使用步驟：
echo    1. 收集驗證碼圖片 (建議100-500張)
echo    2. 手動標註正確答案
echo    3. 生成訓練數據
echo    4. 測試模型效果
echo.
echo 💡 提示：標註時請用圖片檢視器開啟圖片查看
echo.

cd /d "%~dp0"

if not exist "tesseract_trainer.py" (
    echo ❌ 找不到 tesseract_trainer.py 文件
    echo 請確保在正確的目錄中運行此批次文件
    pause
    exit /b 1
)

echo 🚀 啟動訓練器...
echo.

python tesseract_trainer.py

echo.
echo 👋 程式已結束
pause

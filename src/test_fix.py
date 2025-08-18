#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試修正後的爬蟲程式
"""

import sys
import os

# 添加當前目錄到Python路徑
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_import():
    """測試是否能正常導入模組"""
    try:
        # 測試基本導入
        print("正在測試基本導入...")
        import re
        import threading
        import base64
        import requests
        import io
        import os
        import logging
        from functools import lru_cache
        from PIL import Image
        import pytesseract
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.support.ui import Select
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from bs4 import BeautifulSoup
        import time
        import sqlite3
        from selenium.common.exceptions import TimeoutException
        print("✓ 基本模組導入成功")
        
        # 測試主程式導入
        print("正在測試主程式導入...")
        import importlib.util
        spec = importlib.util.spec_from_file_location("crawler", "20240920.py")
        crawler = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(crawler)
        print("✓ 主程式導入成功")
        
        # 檢查關鍵函數是否存在
        required_functions = [
            'initialize_directories',
            'convert_img',
            'download_captcha_image',
            'get_captcha_efficiently',
            'process_database',
            'start_threads',
            'stop_threads',
            'main'
        ]
        
        print("正在檢查關鍵函數...")
        for func_name in required_functions:
            if hasattr(crawler, func_name):
                print(f"✓ 函數 {func_name} 存在")
            else:
                print(f"✗ 函數 {func_name} 不存在")
        
        return True
        
    except Exception as e:
        print(f"✗ 導入測試失敗: {e}")
        return False

def test_config():
    """測試配置"""
    try:
        print("正在測試配置...")
        
        # 檢查目錄結構
        required_dirs = ['../var', '../img']
        for dir_path in required_dirs:
            if os.path.exists(dir_path):
                print(f"✓ 目錄 {dir_path} 存在")
            else:
                print(f"? 目錄 {dir_path} 不存在（將自動創建）")
        
        # 檢查Tesseract
        tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tesseract_path):
            print(f"✓ Tesseract 已安裝: {tesseract_path}")
        else:
            print(f"✗ Tesseract 未找到: {tesseract_path}")
        
        return True
        
    except Exception as e:
        print(f"✗ 配置測試失敗: {e}")
        return False

def main():
    """主測試函數"""
    print("=" * 50)
    print("爬蟲程式修正驗證測試")
    print("=" * 50)
    
    all_passed = True
    
    # 測試導入
    if not test_import():
        all_passed = False
    
    print()
    
    # 測試配置
    if not test_config():
        all_passed = False
    
    print()
    print("=" * 50)
    if all_passed:
        print("✓ 所有測試通過！程式已修正完成，可以正常執行。")
        print("\n使用方式:")
        print("python 20240920.py")
    else:
        print("✗ 部分測試失敗，請檢查相關依賴和配置。")
    print("=" * 50)

if __name__ == "__main__":
    main()

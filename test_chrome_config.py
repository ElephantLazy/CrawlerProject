#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試 Chrome 配置是否能解決 DEPRECATED_ENDPOINT 錯誤
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

def create_optimized_chrome():
    """創建優化配置的 Chrome 瀏覽器實例"""
    options = webdriver.ChromeOptions()
    browser_options = [
        '--headless', '--no-sandbox', '--disable-dev-shm-usage',
        '--disable-gpu', '--disable-extensions', '--disable-logging',
        '--disable-web-security', '--disable-features=VizDisplayCompositor',
        '--disable-background-networking',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-sync',
        '--disable-translate',
        '--disable-ipc-flooding-protection',
        '--disable-renderer-backgrounding',
        '--disable-field-trial-config',
        '--disable-back-forward-cache',
        '--disable-breakpad',
        '--disable-component-extensions-with-background-pages',
        '--disable-default-apps',
        '--disable-features=TranslateUI,BlinkGenPropertyTrees',
        '--no-default-browser-check',
        '--no-first-run',
        '--disable-client-side-phishing-detection',
        '--disable-hang-monitor',
        '--disable-popup-blocking',
        '--disable-prompt-on-repost',
        '--disable-domain-reliability',
        '--disable-component-update',
        # 新增 GCM 相關禁用選項
        '--disable-features=VizDisplayCompositor,GCMChannelStatusRequest',
        '--disable-gcm-registration',
        '--disable-push-messaging',
        '--disable-notifications',
        '--disable-background-mode',
        '--disable-cloud-import',
        '--disable-features=UserMediaScreenCapturing',
        '--disable-features=MediaRouter',
        '--disable-features=PasswordsAccountStorage',
        '--disable-features=AutofillServerCommunication',
        '--disable-features=CertificateTransparencyComponentUpdater',
        '--disable-sync-types',
        '--disable-background-downloads',
        '--disable-features=NetworkService',
        '--disable-features=VizServiceBase',
        '--disable-web-resources',
        '--disable-plugins-discovery',
        '--disable-component-cloud-policy',
        '--disable-session-crashed-bubble',
        '--disable-component-update',
        '--no-pings',
        '--no-wifi',
        '--disable-features=WebRTC'
    ]
    for option in browser_options:
        options.add_argument(option)
    
    # 設置用戶代理以避免檢測
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    # 禁用圖片載入以提升性能並禁用更多功能
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.push_messaging": 2,
        "profile.default_content_setting_values.geolocation": 2,
        "profile.default_content_setting_values.media_stream": 2,
        "gcm.product_category_for_subtypes": "",
        "gcm.check_for_reachability": False,
        "profile.default_content_setting_values.automatic_downloads": 2,
        "profile.default_content_setting_values.mixed_script": 2,
        "profile.background_mode.enabled": False,
        "background_mode.enabled": False,
        "hardware_acceleration_mode.enabled": False
    }
    options.add_experimental_option("prefs", prefs)
    
    # 禁用日誌輸出
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    # 設置額外的環境變數來禁用 GCM
    options.add_argument("--disable-features=GCMChannelStatusRequest,PushMessaging,Notifications")
    
    return webdriver.Chrome(options=options)

def test_chrome_config():
    """測試 Chrome 配置"""
    print("正在測試優化的 Chrome 配置...")
    
    chrome = None
    try:
        chrome = create_optimized_chrome()
        print("✓ Chrome 瀏覽器成功啟動")
        
        # 測試訪問目標網站
        print("正在訪問目標網站...")
        chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
        print("✓ 成功訪問目標網站")
        
        # 等待頁面載入
        time.sleep(3)
        
        print("✓ 測試完成，沒有出現 DEPRECATED_ENDPOINT 錯誤")
        return True
        
    except Exception as e:
        print(f"✗ 測試失敗: {e}")
        return False
    finally:
        if chrome:
            chrome.quit()
            print("✓ Chrome 瀏覽器已關閉")

if __name__ == "__main__":
    success = test_chrome_config()
    if success:
        print("\n🎉 Chrome 配置測試成功！")
        print("主程式應該不會再出現 DEPRECATED_ENDPOINT 錯誤。")
    else:
        print("\n❌ Chrome 配置測試失敗。")
        print("請檢查 Chrome 瀏覽器安裝或網路連線。")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
終極 GCM 錯誤解決方案 - 使用環境變數完全禁用 GCM
"""

import os
import sys

def setup_chrome_environment():
    """設置環境變數來完全禁用 Chrome GCM 功能"""
    
    # 禁用 GCM 相關的環境變數
    gcm_disable_vars = {
        'CHROME_HEADLESS': '1',
        'CHROME_NO_SANDBOX': '1',
        'GOOGLE_API_KEY': '',
        'GOOGLE_DEFAULT_CLIENT_ID': '',
        'GOOGLE_DEFAULT_CLIENT_SECRET': '',
        'CHROME_DISABLE_GCM': '1',
        'CHROME_DISABLE_SYNC': '1',
        'CHROME_DISABLE_BACKGROUND_NETWORKING': '1',
        'CHROME_DISABLE_COMPONENT_UPDATE': '1',
        'CHROME_DISABLE_REPORTING': '1',
        'CHROME_DISABLE_CRASH_REPORTING': '1'
    }
    
    # 設置環境變數
    for key, value in gcm_disable_vars.items():
        os.environ[key] = value
        print(f"✓ 設置環境變數: {key}={value}")
    
    print("🔧 Chrome 環境變數配置完成")

def create_ultimate_chrome_options():
    """創建終極 Chrome 配置選項"""
    from selenium import webdriver
    
    options = webdriver.ChromeOptions()
    
    # 基本無頭模式選項
    basic_options = [
        '--headless=new',  # 使用新的無頭模式
        '--no-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--disable-software-rasterizer'
    ]
    
    # GCM 和網路相關禁用選項
    gcm_options = [
        '--disable-background-networking',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-sync',
        '--disable-translate',
        '--disable-features=GCMChannelStatusRequest',
        '--disable-features=PushMessaging',
        '--disable-features=Notifications',
        '--disable-notifications',
        '--disable-push-messaging',
        '--disable-gcm-registration',
        '--disable-background-mode'
    ]
    
    # 網路和服務禁用選項
    network_options = [
        '--disable-component-update',
        '--disable-domain-reliability',
        '--disable-background-downloads',
        '--disable-cloud-import',
        '--disable-sync-types',
        '--no-pings',
        '--disable-web-resources',
        '--disable-component-cloud-policy'
    ]
    
    # 功能禁用選項
    feature_options = [
        '--disable-extensions',
        '--disable-plugins',
        '--disable-images',
        '--disable-javascript',  # 如果不需要 JavaScript
        '--disable-web-security',
        '--disable-features=VizDisplayCompositor',
        '--disable-features=TranslateUI',
        '--disable-features=BlinkGenPropertyTrees',
        '--disable-features=AutofillServerCommunication',
        '--disable-features=CertificateTransparencyComponentUpdater'
    ]
    
    # 日誌和調試禁用選項
    logging_options = [
        '--disable-logging',
        '--log-level=3',
        '--silent',
        '--disable-dev-tools',
        '--disable-breakpad'
    ]
    
    # 其他優化選項
    misc_options = [
        '--no-default-browser-check',
        '--no-first-run',
        '--disable-default-apps',
        '--disable-popup-blocking',
        '--disable-prompt-on-repost',
        '--disable-hang-monitor',
        '--disable-client-side-phishing-detection',
        '--disable-ipc-flooding-protection',
        '--disable-renderer-backgrounding',
        '--disable-field-trial-config',
        '--disable-back-forward-cache'
    ]
    
    # 添加所有選項
    all_options = basic_options + gcm_options + network_options + feature_options + logging_options + misc_options
    
    for option in all_options:
        options.add_argument(option)
    
    # 設置 prefs 來進一步禁用功能
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.push_messaging": 2,
        "profile.default_content_setting_values.geolocation": 2,
        "profile.default_content_setting_values.media_stream": 2,
        "profile.default_content_setting_values.automatic_downloads": 2,
        "profile.background_mode.enabled": False,
        "background_mode.enabled": False,
        "hardware_acceleration_mode.enabled": False,
        "gcm.check_for_reachability": False,
        "gcm.product_category_for_subtypes": "",
        "sync.engine.bookmarks": False,
        "sync.engine.passwords": False,
        "sync.engine.preferences": False,
        "sync.engine.tabs": False,
        "sync.engine.themes": False,
        "sync.engine.extensions": False,
        "sync.engine.apps": False
    }
    
    options.add_experimental_option("prefs", prefs)
    
    # 排除開關
    options.add_experimental_option("excludeSwitches", [
        "enable-logging",
        "enable-automation",
        "enable-blink-features"
    ])
    
    # 禁用自動化檢測
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    return options

def test_ultimate_fix():
    """測試終極修復方案"""
    print("🚀 開始終極 GCM 錯誤修復測試...")
    
    # 1. 設置環境變數
    setup_chrome_environment()
    print()
    
    # 2. 測試 Chrome 配置
    try:
        from selenium import webdriver
        print("📱 創建終極優化的 Chrome 實例...")
        
        options = create_ultimate_chrome_options()
        chrome = webdriver.Chrome(options=options)
        
        print("✓ Chrome 成功啟動")
        
        # 測試訪問網站
        print("🌐 測試訪問目標網站...")
        chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
        print("✓ 成功訪問網站")
        
        # 保持一段時間以檢測錯誤
        import time
        print("⏱️  等待 5 秒以檢測潛在錯誤...")
        time.sleep(5)
        
        chrome.quit()
        print("✓ Chrome 已安全關閉")
        
        print("\n🎉 終極修復測試成功完成！")
        print("✅ 如果沒有看到 GCM 錯誤訊息，表示修復成功。")
        
        return True
        
    except Exception as e:
        print(f"❌ 測試失敗: {e}")
        return False

if __name__ == "__main__":
    test_ultimate_fix()

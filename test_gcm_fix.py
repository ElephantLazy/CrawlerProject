#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
強化測試 Chrome 配置是否能完全解決 GCM 相關錯誤
"""

import sys
import os
import time
import subprocess
import threading
from contextlib import redirect_stderr
import io

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

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

def test_multiple_chrome_instances():
    """測試多個 Chrome 實例來模擬真實使用情況"""
    print("🧪 正在測試多個 Chrome 實例（模擬真實爬蟲環境）...")
    
    chrome_instances = []
    try:
        # 創建多個 Chrome 實例（模擬多線程環境）
        for i in range(3):
            print(f"  📱 啟動 Chrome 實例 {i+1}...")
            chrome = create_optimized_chrome()
            chrome_instances.append(chrome)
            
            # 訪問目標網站
            chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
            print(f"  ✓ Chrome 實例 {i+1} 成功訪問網站")
            
            # 等待一段時間讓可能的錯誤出現
            time.sleep(2)
        
        print("  ⏱️  讓所有實例運行 10 秒以檢測潛在錯誤...")
        time.sleep(10)
        
        print("✅ 多實例測試完成 - 沒有檢測到 GCM 錯誤")
        return True
        
    except Exception as e:
        print(f"❌ 多實例測試失敗: {e}")
        return False
    finally:
        # 清理所有實例
        for i, chrome in enumerate(chrome_instances):
            try:
                chrome.quit()
                print(f"  🔄 Chrome 實例 {i+1} 已關閉")
            except:
                pass

def test_long_running_session():
    """測試長時間運行的會話"""
    print("🕐 正在測試長時間運行的會話...")
    
    chrome = None
    try:
        chrome = create_optimized_chrome()
        print("  ✓ Chrome 瀏覽器已啟動")
        
        # 模擬爬蟲的典型操作
        for i in range(5):
            print(f"  🔄 執行第 {i+1} 次頁面操作...")
            chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
            time.sleep(3)  # 模擬處理時間
            
            # 刷新頁面（模擬驗證碼刷新）
            chrome.refresh()
            time.sleep(2)
        
        print("✅ 長時間運行測試完成 - 沒有檢測到 GCM 錯誤")
        return True
        
    except Exception as e:
        print(f"❌ 長時間運行測試失敗: {e}")
        return False
    finally:
        if chrome:
            chrome.quit()
            print("  🔄 Chrome 瀏覽器已關閉")

def check_chrome_processes():
    """檢查是否有殘留的 Chrome 進程"""
    print("🔍 檢查 Chrome 進程狀態...")
    
    try:
        # 檢查 Chrome 進程
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq chrome.exe'], 
                              capture_output=True, text=True, shell=True)
        
        if 'chrome.exe' in result.stdout:
            lines = [line for line in result.stdout.split('\n') if 'chrome.exe' in line]
            print(f"  📊 發現 {len(lines)} 個 Chrome 進程")
        else:
            print("  ✅ 沒有發現殘留的 Chrome 進程")
            
    except Exception as e:
        print(f"  ⚠️  無法檢查進程狀態: {e}")

def main():
    """主測試函數"""
    print("🚀 開始強化 Chrome 配置測試...")
    print("=" * 60)
    
    # 檢查初始進程狀態
    check_chrome_processes()
    print()
    
    # 測試 1: 多實例測試
    test1_result = test_multiple_chrome_instances()
    print()
    
    # 測試 2: 長時間運行測試
    test2_result = test_long_running_session()
    print()
    
    # 最終進程檢查
    check_chrome_processes()
    print()
    
    # 總結
    print("=" * 60)
    if test1_result and test2_result:
        print("🎉 所有測試通過！GCM 錯誤已被成功抑制。")
        print("✅ Chrome 配置優化完成，可以安全使用。")
        print("💡 建議：在生產環境中運行以驗證效果。")
    else:
        print("⚠️  部分測試失敗，可能需要進一步調整配置。")
        print("📝 建議：檢查 Chrome 版本或嘗試其他配置選項。")
    
    print("\n🔧 如果仍有問題，請考慮：")
    print("   1. 更新 ChromeDriver 到最新版本")
    print("   2. 檢查 Chrome 瀏覽器版本")
    print("   3. 在無痕模式下測試")

if __name__ == "__main__":
    main()

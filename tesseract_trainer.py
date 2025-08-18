#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tesseract-OCR 驗證碼訓練數據生成器

此工具用於：
1. 下載驗證碼圖片
2. 手動標註正確答案
3. 生成 Tesseract 訓練數據
4. 創建自定義模型

使用方法：
1. 運行腳本收集驗證碼圖片
2. 手動標註每張圖片的正確答案
3. 生成訓練數據集
4. 訓練自定義 Tesseract 模型
"""

import os
import sys
import time
import json
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

# 添加主程式路徑以導入相關模組
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 圖片處理
from PIL import Image, ImageDraw, ImageFont
import pytesseract

# 網頁爬取
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import base64
import io

class TesseractTrainer:
    def __init__(self):
        self.base_dir = Path("./tesseract_training")
        self.images_dir = self.base_dir / "images"
        self.annotations_dir = self.base_dir / "annotations"
        self.training_dir = self.base_dir / "training_data"
        self.models_dir = self.base_dir / "models"
        
        # 創建必要目錄
        for dir_path in [self.images_dir, self.annotations_dir, self.training_dir, self.models_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)
        
        # 標註數據文件
        self.annotations_file = self.base_dir / "annotations.json"
        self.load_annotations()
        
        print(f"✅ 訓練器初始化完成")
        print(f"📁 工作目錄: {self.base_dir}")
        print(f"🖼️  圖片目錄: {self.images_dir}")
        print(f"📝 標註目錄: {self.annotations_dir}")

    def load_annotations(self):
        """載入已有的標註數據"""
        if self.annotations_file.exists():
            with open(self.annotations_file, 'r', encoding='utf-8') as f:
                self.annotations = json.load(f)
        else:
            self.annotations = {}
        print(f"📊 已載入 {len(self.annotations)} 個標註")

    def save_annotations(self):
        """保存標註數據"""
        with open(self.annotations_file, 'w', encoding='utf-8') as f:
            json.dump(self.annotations, f, ensure_ascii=False, indent=2)
        print(f"💾 已保存 {len(self.annotations)} 個標註")

    def create_optimized_chrome(self):
        """創建優化的 Chrome 瀏覽器實例（從主程式複製）"""
        options = webdriver.ChromeOptions()
        
        # 基本選項
        core_options = [
            '--headless=new',
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--disable-software-rasterizer',
            '--disable-background-networking',
            '--disable-sync',
            '--disable-logging',
            '--log-level=3',
            '--silent'
        ]
        
        for option in core_options:
            options.add_argument(option)
        
        # 設置用戶代理
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        return webdriver.Chrome(options=options)

    def download_captcha_image(self, chrome, index):
        """下載驗證碼圖片（從主程式修改）"""
        try:
            # 點擊刷新驗證碼按鈕
            element = WebDriverWait(chrome, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="queryForm"]/div[1]/div[2]/div[2]/div/div/etw-captcha/div/button[1]'))
            )
            element.click()
            time.sleep(1)  # 等待驗證碼載入
            
            # 嘗試不同的驗證碼圖片選擇器
            captcha_selectors = [
                'img[src*="captcha"]',
                'img[alt*="captcha"]',
                'img[id*="captcha"]',
                '#queryForm img',
                'etw-captcha img'
            ]
            
            captcha_img = None
            for selector in captcha_selectors:
                try:
                    captcha_img = chrome.find_element(By.CSS_SELECTOR, selector)
                    if captcha_img:
                        break
                except:
                    continue
            
            if not captcha_img:
                print(f"❌ 無法找到驗證碼圖片元素")
                return None
                
            # 獲取圖片的 src 屬性
            img_src = captcha_img.get_attribute('src')
            
            if img_src.startswith('data:image'):
                # 如果是 base64 編碼的圖片
                header, encoded = img_src.split(',', 1)
                img_data = base64.b64decode(encoded)
                img = Image.open(io.BytesIO(img_data))
            else:
                # 如果是 URL 連結
                response = requests.get(img_src, stream=True)
                if response.status_code == 200:
                    img = Image.open(io.BytesIO(response.content))
                else:
                    print(f"❌ 無法下載驗證碼圖片，狀態碼: {response.status_code}")
                    return None
            
            # 保存原始圖片
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"captcha_{timestamp}_{index:04d}.png"
            img_path = self.images_dir / filename
            img.save(img_path)
            
            print(f"✅ 成功下載驗證碼: {filename}")
            return img_path, img
            
        except Exception as e:
            print(f"❌ 下載驗證碼失敗: {e}")
            return None

    def collect_captcha_images(self, num_images=100):
        """收集驗證碼圖片"""
        print(f"🚀 開始收集 {num_images} 張驗證碼圖片...")
        
        chrome = None
        collected = 0
        
        try:
            chrome = self.create_optimized_chrome()
            print("✅ Chrome 瀏覽器已啟動")
            
            while collected < num_images:
                try:
                    # 訪問驗證碼頁面
                    chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
                    time.sleep(2)
                    
                    # 下載驗證碼圖片
                    result = self.download_captcha_image(chrome, collected + 1)
                    
                    if result:
                        collected += 1
                        print(f"📸 已收集 {collected}/{num_images} 張圖片")
                        
                        # 每10張圖片休息一下
                        if collected % 10 == 0:
                            print("⏸️  休息 3 秒...")
                            time.sleep(3)
                    else:
                        print("⚠️  圖片下載失敗，重試中...")
                        time.sleep(1)
                        
                except Exception as e:
                    print(f"❌ 收集過程中發生錯誤: {e}")
                    time.sleep(2)
                    continue
                    
        except Exception as e:
            print(f"❌ 收集失敗: {e}")
        finally:
            if chrome:
                chrome.quit()
                print("🔄 Chrome 瀏覽器已關閉")
        
        print(f"🎉 收集完成！總共收集了 {collected} 張驗證碼圖片")
        return collected

    def manual_annotation_mode(self):
        """手動標註模式"""
        print("📝 進入手動標註模式...")
        print("💡 提示：")
        print("   - 輸入驗證碼的正確答案")
        print("   - 輸入 'skip' 跳過該圖片")
        print("   - 輸入 'quit' 退出標註")
        print("   - 輸入 'save' 保存進度")
        print("-" * 50)
        
        # 獲取所有未標註的圖片
        image_files = list(self.images_dir.glob("*.png"))
        unannotated = [f for f in image_files if f.name not in self.annotations]
        
        print(f"📊 發現 {len(image_files)} 張圖片，其中 {len(unannotated)} 張未標註")
        
        if not unannotated:
            print("✅ 所有圖片都已標註完成！")
            return
        
        annotated_count = 0
        
        for img_path in unannotated:
            try:
                # 顯示圖片信息
                print(f"\n📸 當前圖片: {img_path.name}")
                
                # 嘗試用當前 Tesseract 識別（作為參考）
                try:
                    img = Image.open(img_path)
                    current_ocr = pytesseract.image_to_string(img, config='--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz')
                    current_ocr = ''.join(c for c in current_ocr if c.isalnum())
                    if current_ocr:
                        print(f"🤖 當前OCR識別: '{current_ocr}' (僅供參考)")
                    else:
                        print("🤖 當前OCR無法識別")
                except:
                    print("🤖 當前OCR識別失敗")
                
                # 提示用戶開啟圖片
                print(f"📁 請開啟圖片查看: {img_path}")
                print("   (可以用 Windows 照片檢視器或其他圖片軟體開啟)")
                
                # 獲取用戶輸入
                user_input = input("✏️  請輸入正確的驗證碼 (6位英數字): ").strip()
                
                if user_input.lower() == 'quit':
                    print("👋 退出標註模式")
                    break
                elif user_input.lower() == 'skip':
                    print("⏭️  跳過此圖片")
                    continue
                elif user_input.lower() == 'save':
                    self.save_annotations()
                    print("💾 進度已保存")
                    continue
                elif len(user_input) == 6 and user_input.isalnum():
                    # 保存標註
                    self.annotations[img_path.name] = {
                        'text': user_input.upper(),
                        'timestamp': datetime.now().isoformat(),
                        'image_path': str(img_path)
                    }
                    annotated_count += 1
                    print(f"✅ 標註保存: {user_input.upper()}")
                    
                    # 每標註10張自動保存
                    if annotated_count % 10 == 0:
                        self.save_annotations()
                        print(f"💾 自動保存進度 ({annotated_count} 張已標註)")
                else:
                    print("❌ 請輸入6位英數字驗證碼")
                    # 重新處理同一張圖片
                    continue
                    
            except KeyboardInterrupt:
                print("\n⏸️  用戶中斷，保存進度...")
                self.save_annotations()
                break
            except Exception as e:
                print(f"❌ 標註過程中發生錯誤: {e}")
                continue
        
        # 最終保存
        self.save_annotations()
        print(f"\n🎉 標註完成！本次標註了 {annotated_count} 張圖片")
        print(f"📊 總標註數量: {len(self.annotations)}")

    def create_training_data(self):
        """創建 Tesseract 訓練數據"""
        print("🏗️  創建 Tesseract 訓練數據...")
        
        if not self.annotations:
            print("❌ 沒有標註數據，請先進行手動標註")
            return False
        
        # 清理舊的訓練數據
        if self.training_dir.exists():
            shutil.rmtree(self.training_dir)
        self.training_dir.mkdir(parents=True, exist_ok=True)
        
        # 創建字體文件（如果需要）
        box_files = []
        tif_files = []
        
        for i, (filename, annotation) in enumerate(self.annotations.items()):
            try:
                img_path = self.images_dir / filename
                if not img_path.exists():
                    print(f"⚠️  圖片不存在: {filename}")
                    continue
                
                # 載入圖片
                img = Image.open(img_path)
                
                # 轉換為灰度並調整大小（Tesseract 訓練需要）
                if img.mode != 'L':
                    img = img.convert('L')
                
                # 放大圖片以提高訓練效果
                width, height = img.size
                img = img.resize((width * 2, height * 2), Image.Resampling.LANCZOS)
                
                # 保存為 TIF 格式（Tesseract 訓練格式）
                base_name = f"captcha.font.exp{i:04d}"
                tif_path = self.training_dir / f"{base_name}.tif"
                img.save(tif_path)
                tif_files.append(tif_path)
                
                # 創建 BOX 文件（字符位置信息）
                box_path = self.training_dir / f"{base_name}.box"
                self.create_box_file(img, annotation['text'], box_path)
                box_files.append(box_path)
                
                print(f"✅ 處理完成: {filename} -> {base_name}")
                
            except Exception as e:
                print(f"❌ 處理失敗 {filename}: {e}")
                continue
        
        print(f"🎉 訓練數據創建完成！")
        print(f"📊 生成了 {len(tif_files)} 個 TIF 文件和 {len(box_files)} 個 BOX 文件")
        
        # 創建訓練腳本
        self.create_training_script()
        
        return True

    def create_box_file(self, img, text, box_path):
        """創建 BOX 文件（簡化版本）"""
        width, height = img.size
        char_width = width // len(text)
        
        with open(box_path, 'w', encoding='utf-8') as f:
            for i, char in enumerate(text):
                # 簡單的字符邊界估算
                left = i * char_width
                right = (i + 1) * char_width
                top = height - 5  # 從底部往上
                bottom = 5  # 從頂部往下
                
                # BOX 格式: 字符 左 下 右 上 頁碼
                f.write(f"{char} {left} {bottom} {right} {top} 0\n")

    def create_training_script(self):
        """創建 Tesseract 訓練腳本"""
        script_content = '''@echo off
REM Tesseract 訓練腳本
REM 請確保已安裝 Tesseract 和訓練工具

echo 開始 Tesseract 訓練...

REM 設置語言名稱
set LANG_NAME=captcha

REM 合併所有 BOX 和 TIF 文件
echo 合併訓練文件...
copy /b *.box %LANG_NAME%.box
for %%f in (*.tif) do (
    echo 處理 %%f...
    tesseract %%f %%~nf batch.nochop makebox
)

REM 生成訓練文件
echo 生成訓練文件...
tesseract %LANG_NAME%.font.exp0000.tif %LANG_NAME%.font.exp0000 batch.nochop makebox
for %%f in (*.tif) do tesseract %%f %%~nf box.train

REM 計算字符集
echo 計算字符集...
unicharset_extractor *.box

REM 生成字形統計
echo 生成字形統計...
shapeclustering -F font_properties -U unicharset *.tr

REM 生成聚類
echo 生成聚類...
mftraining -F font_properties -U unicharset -O %LANG_NAME%.unicharset *.tr

REM 生成標準化特徵文件
echo 生成標準化特徵文件...
cntraining *.tr

REM 重命名文件
echo 重命名輸出文件...
ren normproto %LANG_NAME%.normproto
ren inttemp %LANG_NAME%.inttemp
ren pffmtable %LANG_NAME%.pffmtable
ren shapetable %LANG_NAME%.shapetable

REM 合併訓練文件
echo 合併最終模型...
combine_tessdata %LANG_NAME%.

echo 訓練完成！生成的模型文件: %LANG_NAME%.traineddata
echo 請將此文件複製到 Tesseract 的 tessdata 目錄中
pause
'''
        
        script_path = self.training_dir / "train_model.bat"
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        print(f"📝 訓練腳本已生成: {script_path}")
        print("💡 運行此腳本來訓練自定義模型（需要安裝 Tesseract 訓練工具）")

    def show_statistics(self):
        """顯示統計信息"""
        print("\n📊 訓練數據統計:")
        print("-" * 40)
        
        # 圖片統計
        image_files = list(self.images_dir.glob("*.png"))
        print(f"🖼️  總圖片數量: {len(image_files)}")
        print(f"📝 已標註數量: {len(self.annotations)}")
        print(f"❓ 未標註數量: {len(image_files) - len(self.annotations)}")
        
        if self.annotations:
            # 字符統計
            all_chars = ''.join(ann['text'] for ann in self.annotations.values())
            char_count = {}
            for char in all_chars:
                char_count[char] = char_count.get(char, 0) + 1
            
            print(f"\n🔤 字符分佈:")
            for char, count in sorted(char_count.items()):
                print(f"   {char}: {count} 次")
            
            # 長度統計
            lengths = [len(ann['text']) for ann in self.annotations.values()]
            print(f"\n📏 驗證碼長度:")
            print(f"   平均長度: {sum(lengths)/len(lengths):.1f}")
            print(f"   最短: {min(lengths)}, 最長: {max(lengths)}")

    def interactive_menu(self):
        """交互式菜單"""
        while True:
            print("\n" + "="*50)
            print("🎯 Tesseract-OCR 驗證碼訓練器")
            print("="*50)
            print("1. 📸 收集驗證碼圖片")
            print("2. 📝 手動標註驗證碼")
            print("3. 🏗️  生成訓練數據")
            print("4. 📊 查看統計信息")
            print("5. 🧪 測試當前模型")
            print("6. 🚪 退出")
            print("-"*50)
            
            choice = input("請選擇操作 (1-6): ").strip()
            
            if choice == '1':
                num = input("請輸入要收集的圖片數量 (預設100): ").strip()
                try:
                    num = int(num) if num else 100
                    self.collect_captcha_images(num)
                except ValueError:
                    print("❌ 請輸入有效數字")
                    
            elif choice == '2':
                self.manual_annotation_mode()
                
            elif choice == '3':
                self.create_training_data()
                
            elif choice == '4':
                self.show_statistics()
                
            elif choice == '5':
                self.test_current_model()
                
            elif choice == '6':
                print("👋 再見！")
                break
                
            else:
                print("❌ 無效選擇，請重新輸入")

    def test_current_model(self):
        """測試當前模型效果"""
        print("🧪 測試當前 Tesseract 模型效果...")
        
        if not self.annotations:
            print("❌ 沒有標註數據用於測試")
            return
        
        correct = 0
        total = 0
        
        for filename, annotation in list(self.annotations.items())[:10]:  # 測試前10張
            try:
                img_path = self.images_dir / filename
                if not img_path.exists():
                    continue
                
                img = Image.open(img_path)
                
                # 使用當前 Tesseract 識別
                ocr_result = pytesseract.image_to_string(
                    img, 
                    config='--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
                )
                ocr_result = ''.join(c for c in ocr_result if c.isalnum()).upper()
                
                expected = annotation['text'].upper()
                total += 1
                
                if ocr_result == expected:
                    correct += 1
                    print(f"✅ {filename}: '{ocr_result}' == '{expected}'")
                else:
                    print(f"❌ {filename}: '{ocr_result}' != '{expected}'")
                    
            except Exception as e:
                print(f"⚠️  測試失敗 {filename}: {e}")
        
        if total > 0:
            accuracy = correct / total * 100
            print(f"\n📊 測試結果: {correct}/{total} 正確，準確率: {accuracy:.1f}%")
        else:
            print("❌ 沒有可測試的數據")


def main():
    """主程式入口"""
    print("🎯 Tesseract-OCR 驗證碼訓練器啟動中...")
    
    try:
        trainer = TesseractTrainer()
        trainer.interactive_menu()
    except KeyboardInterrupt:
        print("\n👋 程式已中斷")
    except Exception as e:
        print(f"❌ 程式執行錯誤: {e}")


if __name__ == "__main__":
    main()

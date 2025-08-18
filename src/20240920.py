import re
import threading
import base64
import requests
import io
import os
import logging
from functools import lru_cache
# 圖片處理
from PIL import Image
# 文字識別
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

# 導入優化模組
try:
    from performance_config import PerformanceConfig
    from performance_monitor import monitor, start_monitoring
except ImportError:
    # 如果模組不存在，使用預設配置
    class PerformanceConfig:
        MAX_THREADS = 6
        BATCH_SIZE = 20
        CAPTCHA_RETRY_MAX = 2
        ELEMENT_WAIT_TIME = 5
        CAPTCHA_REFRESH_WAIT = 0.3
        THREAD_START_DELAY = 0.1
        TESSERACT_CONFIG = '--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
        IMAGE_SCALE_FACTOR = 3
        BINARY_THRESHOLD = 140
    
    class DummyMonitor:
        def record_processed(self): pass
        def record_success(self): pass
        def record_error(self): pass
        def print_stats(self): pass
    
    monitor = DummyMonitor()
    def start_monitoring(): pass

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

#全域變數
left =  615
top = 600
right =  793
bottom = 681
# 二值化閾值
threshold = 150

# 初始化函數
def initialize_directories():
    """創建必要的目錄"""
    if not os.path.exists('./img'):
        os.makedirs('./img')
        print("已創建 img 目錄")

#影像二值化（優化版本）
@lru_cache(maxsize=32)
def convert_img_cached(img_hash, threshold):
    """快取版本的圖像二值化"""
    # 這是一個概念性的實現，實際使用時需要處理圖像序列化
    pass

def convert_img(img, threshold):
    """優化的圖像二值化"""
    if img.mode != "L":
        img = img.convert("L")
    
    # 使用numpy加速處理（如果可用）
    try:
        import numpy as np
        img_array = np.array(img)
        img_array = np.where(img_array > threshold, 255, 0).astype(np.uint8)
        return Image.fromarray(img_array)
    except ImportError:
        # 回退到原始方法
        pixels = img.load()
        width, height = img.size
        for x in range(width):
            for y in range(height):
                if pixels[x, y] > threshold:
                    pixels[x, y] = 255
                else:
                    pixels[x, y] = 0
        return img

#降躁
def clearImg(img):
    data = img.getdata()
    w, h = img.size
    count = 0
    for x in range(1, h-1):
        for y in range(1, h - 1):
            # 找出各個像素方向
            mid_pixel = data[w * y + x]
            if mid_pixel == 0:
                top_pixel = data[w * (y - 1) + x]
                left_pixel = data[w * y + (x - 1)]
                down_pixel = data[w * (y + 1) + x]
                right_pixel = data[w * y + (x + 1)]
                if top_pixel == 0:
                    count += 1
                    if left_pixel == 0:
                        count += 1
                        if down_pixel == 0:
                            count += 1
                            if right_pixel == 0:
                                count += 1
                                if count > 4:
                                    img.putpixel((x, y), 0)
    return img

# 自適應閾值二值化（專為Tesseract優化）
def convert_img_with_adaptive_threshold(img):
    try:
        # 計算Otsu閾值
        pixels = list(img.getdata())
        histogram = [0] * 256
        for pixel in pixels:
            histogram[pixel] += 1
        
        total_pixels = len(pixels)
        sum_total = sum(i * histogram[i] for i in range(256))
        
        sum_background = 0
        weight_background = 0
        max_variance = 0
        best_threshold = 0
        
        for t in range(256):
            weight_background += histogram[t]
            if weight_background == 0:
                continue
                
            weight_foreground = total_pixels - weight_background
            if weight_foreground == 0:
                break
                
            sum_background += t * histogram[t]
            
            if weight_background > 0 and weight_foreground > 0:
                mean_background = sum_background / weight_background
                mean_foreground = (sum_total - sum_background) / weight_foreground
                
                variance_between = weight_background * weight_foreground * (mean_background - mean_foreground) ** 2
                
                if variance_between > max_variance:
                    max_variance = variance_between
                    best_threshold = t
        
        print(f"計算出的Otsu閾值: {best_threshold}")
        
        # 應用閾值
        img = img.convert("L")
        pixels = img.load()
        width, height = img.size
        
        for x in range(width):
            for y in range(height):
                if pixels[x, y] > best_threshold:
                    pixels[x, y] = 255  # 白色背景
                else:
                    pixels[x, y] = 0    # 黑色文字
        
        return img
    except Exception as e:
        print(f"自適應閾值處理失敗: {e}")
        return convert_img(img, threshold)

# 簡化的形態學清理（專為Tesseract設計）
def morphological_cleanup_simple(img):
    try:
        width, height = img.size
        pixels = img.load()
        
        # 創建結果圖像
        result = Image.new('L', (width, height), 255)
        result_pixels = result.load()
        
        # 複製原圖
        for x in range(width):
            for y in range(height):
                result_pixels[x, y] = pixels[x, y]
        
        # 去除孤立像素（開運算的簡化版本）
        for x in range(1, width-1):
            for y in range(1, height-1):
                if pixels[x, y] == 0:  # 黑色像素
                    # 檢查4鄰域
                    neighbors = [
                        pixels[x-1, y], pixels[x+1, y],
                        pixels[x, y-1], pixels[x, y+1]
                    ]
                    # 如果周圍都是白色，則認為是噪點
                    if all(n == 255 for n in neighbors):
                        result_pixels[x, y] = 255
        
        return result
    except Exception as e:
        print(f"形態學清理失敗: {e}")
        return img

# 針對Tesseract的高級降噪
def advanced_denoise_tesseract(img):
    try:
        width, height = img.size
        pixels = img.load()
        
        # 創建結果圖像
        result = Image.new('L', (width, height), 255)
        result_pixels = result.load()
        
        # 複製原圖
        for x in range(width):
            for y in range(height):
                result_pixels[x, y] = pixels[x, y]
        
        # 連通域分析和小區域移除
        for x in range(1, width-1):
            for y in range(1, height-1):
                if pixels[x, y] == 0:  # 黑色像素
                    # 檢查3x3鄰域
                    black_count = 0
                    for dx in [-1, 0, 1]:
                        for dy in [-1, 0, 1]:
                            if x+dx >= 0 and x+dx < width and y+dy >= 0 and y+dy < height:
                                if pixels[x+dx, y+dy] == 0:
                                    black_count += 1
                    
                    # 如果鄰域內黑色像素太少，視為噪點
                    if black_count < 3:
                        result_pixels[x, y] = 255
        
        return result
    except Exception as e:
        print(f"高級降噪失敗: {e}")
        return clearImg(img)

# 為Tesseract添加邊框
def add_tesseract_padding(img):
    try:
        # Tesseract需要圖像周圍有空白邊框
        padding = 15
        width, height = img.size
        
        # 創建新圖像，周圍添加白色邊框
        padded = Image.new('L', (width + 2*padding, height + 2*padding), 255)
        padded.paste(img, (padding, padding))
        
        return padded
    except Exception as e:
        print(f"添加邊框失敗: {e}")
        return img

# 顏色標準化處理
def normalize_captcha_colors(img):
    try:
        from PIL import ImageEnhance
        
        # 增強亮度
        enhancer = ImageEnhance.Brightness(img)
        img = enhancer.enhance(1.2)
        
        # 增強對比度
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.8)
        
        return img
    except Exception as e:
        print(f"顏色標準化失敗: {e}")
        return img

# 多閾值二值化策略
def multi_threshold_binarization(img, i):
    try:
        # 嘗試多個不同的閾值
        thresholds = [120, 140, 160, 180]
        best_img = None
        best_score = 0
        
        for threshold_val in thresholds:
            # 創建二值化圖像
            test_img = img.copy()
            pixels = test_img.load()
            width, height = test_img.size
            
            for x in range(width):
                for y in range(height):
                    if pixels[x, y] > threshold_val:
                        pixels[x, y] = 255
                    else:
                        pixels[x, y] = 0
            
            # 評估這個閾值的效果
            score = evaluate_binarization_quality(test_img)
            print(f"閾值 {threshold_val} 評分: {score}")
            
            if score > best_score:
                best_score = score
                best_img = test_img.copy()
        
        # 如果所有閾值都不好，使用Otsu方法
        if best_img is None:
            best_img = convert_img_with_adaptive_threshold(img)
        
        return best_img
    except Exception as e:
        print(f"多閾值二值化失敗: {e}")
        return convert_img(img, 150)

# 評估二值化品質
def evaluate_binarization_quality(img):
    try:
        pixels = list(img.getdata())
        width, height = img.size
        
        # 計算邊緣數量（邊緣越多，字符越清晰）
        edge_count = 0
        for y in range(1, height-1):
            for x in range(1, width-1):
                current = pixels[y * width + x]
                neighbors = [
                    pixels[(y-1) * width + x],
                    pixels[(y+1) * width + x],
                    pixels[y * width + (x-1)],
                    pixels[y * width + (x+1)]
                ]
                # 如果像素與鄰居有顯著差異，則為邊緣
                if any(abs(current - neighbor) > 100 for neighbor in neighbors):
                    edge_count += 1
        
        # 計算黑白比例（理想的驗證碼應該有適當的黑白比例）
        black_pixels = sum(1 for p in pixels if p < 128)
        white_pixels = len(pixels) - black_pixels
        ratio = min(black_pixels, white_pixels) / max(black_pixels, white_pixels) if max(black_pixels, white_pixels) > 0 else 0
        
        # 綜合評分
        edge_score = edge_count / (width * height) * 100
        ratio_score = ratio * 50
        
        return edge_score + ratio_score
    except:
        return 0

# 字符分離處理
def character_isolation(img, i):
    try:
        width, height = img.size
        pixels = img.load()
        
        # 垂直投影分析
        vertical_projection = []
        for x in range(width):
            black_count = 0
            for y in range(height):
                if pixels[x, y] == 0:  # 黑色像素
                    black_count += 1
            vertical_projection.append(black_count)
        
        # 找出字符邊界
        char_boundaries = []
        in_char = False
        start_x = 0
        
        for x in range(width):
            if vertical_projection[x] > height * 0.1:  # 有足夠的黑色像素
                if not in_char:
                    start_x = x
                    in_char = True
            else:
                if in_char:
                    char_boundaries.append((start_x, x))
                    in_char = False
        
        if in_char:  # 處理最後一個字符
            char_boundaries.append((start_x, width))
        
        print(f"檢測到 {len(char_boundaries)} 個字符邊界: {char_boundaries}")
        
        # 如果字符粘連，嘗試分離
        if len(char_boundaries) < 6:
            img = separate_connected_characters(img, char_boundaries)
        
        return img
    except Exception as e:
        print(f"字符分離失敗: {e}")
        return img

# 分離粘連字符
def separate_connected_characters(img, boundaries):
    try:
        width, height = img.size
        pixels = img.load()
        
        # 對於寬度過大的字符區域，嘗試分離
        for start_x, end_x in boundaries:
            char_width = end_x - start_x
            if char_width > width / 4:  # 如果單個字符太寬，可能是粘連
                # 在中間位置嘗試分離
                mid_x = start_x + char_width // 2
                
                # 尋找最細的連接點
                min_black_count = height
                best_cut_x = mid_x
                
                for x in range(start_x + char_width//4, start_x + 3*char_width//4):
                    black_count = 0
                    for y in range(height):
                        if pixels[x, y] == 0:
                            black_count += 1
                    
                    if black_count < min_black_count:
                        min_black_count = black_count
                        best_cut_x = x
                
                # 如果找到較細的連接點，進行分離
                if min_black_count < height * 0.3:
                    for y in range(height):
                        pixels[best_cut_x, y] = 255  # 白色分離線
        
        return img
    except Exception as e:
        print(f"分離粘連字符失敗: {e}")
        return img

# 最終清理
def final_cleanup_for_tesseract(img):
    try:
        width, height = img.size
        
        # 添加適當的邊框
        padding = 10
        padded = Image.new('L', (width + 2*padding, height + 2*padding), 255)
        padded.paste(img, (padding, padding))
        
        # 確保圖像大小適合Tesseract
        final_width = max(padded.width, 200)  # 最小寬度
        final_height = max(padded.height, 50)  # 最小高度
        
        if padded.width < final_width or padded.height < final_height:
            final_img = Image.new('L', (final_width, final_height), 255)
            paste_x = (final_width - padded.width) // 2
            paste_y = (final_height - padded.height) // 2
            final_img.paste(padded, (paste_x, paste_y))
            return final_img
        
        return padded
    except Exception as e:
        print(f"最終清理失敗: {e}")
        return img

# 簡單增強備用方案
def simple_enhance_fallback(img, i):
    try:
        if img is None:
            return None
        
        # 簡單的處理流程
        if img.mode != 'L':
            img = img.convert('L')
        
        # 放大2倍
        width, height = img.size
        img = img.resize((width * 2, height * 2), Image.Resampling.LANCZOS)
        
        # 基本二值化
        img = convert_img(img, 140)
        
        # 基本降噪
        img = clearImg(img)
        
        # 添加邊框
        padding = 10
        padded = Image.new('L', (img.width + 2*padding, img.height + 2*padding), 255)
        padded.paste(img, (padding, padding))
        
        padded.save(f"./img/captcha_fallback{i}.png")
        return padded
    except Exception as e:
        print(f"備用方案失敗: {e}")
        return None

# 智能Tesseract OCR識別
def smart_tesseract_recognition(img, i):
    try:
        if img is None:
            return ''
        
        # 更實用的Tesseract配置
        tesseract_configs = [
            # 基本配置 - 單詞模式
            '--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz',
            # 更寬鬆的配置
            '--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz',
            # 最簡配置
            '--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz',
            # 無限制配置（作為備用）
            '--psm 8',
        ]
        
        results = []
        
        for idx, config in enumerate(tesseract_configs):
            try:
                result = pytesseract.image_to_string(img, config=config)
                # 清理結果，只保留英數字符
                clean_result = re.sub(r'[^a-zA-Z0-9]', '', result)
                if clean_result and len(clean_result) >= 4:  # 至少4個字符才考慮
                    results.append(clean_result)
                    print(f"配置 {idx+1}: '{clean_result}' (長度: {len(clean_result)})")
            except Exception as e:
                print(f"配置 {idx+1} 失敗: {e}")
                continue
        
        if not results:
            print("所有OCR配置都失敗，嘗試基本識別")
            try:
                result = pytesseract.image_to_string(img)
                clean_result = re.sub(r'[^a-zA-Z0-9]', '', result)
                if clean_result:
                    results.append(clean_result)
            except:
                pass
        
        if not results:
            return ''
        
        # 智能選擇最佳結果
        print(f"所有識別結果: {results}")
        
        # 1. 優先選擇長度為6的結果
        six_char_results = [r for r in results if len(r) == 6]
        if six_char_results:
            print(f"找到6位結果: {six_char_results}")
            return six_char_results[0]  # 取第一個6位結果
        
        # 2. 選擇長度最接近6的結果
        results_by_length = sorted(results, key=lambda x: abs(len(x) - 6))
        best_result = results_by_length[0]
        
        # 3. 如果結果太短，嘗試組合
        if len(best_result) < 6:
            # 嘗試從所有結果中找最長的
            longest_results = sorted(results, key=len, reverse=True)
            if longest_results:
                best_result = longest_results[0]
        
        print(f"選擇的最佳結果: '{best_result}'")
        return best_result
        
    except Exception as e:
        print(f"智能Tesseract識別失敗: {e}")
        return ''

# 驗證碼結果後處理（針對常見錯誤）
def post_process_tesseract_result(result):
    """
    修正Tesseract常見的識別錯誤
    """
    if not result:
        return result
    
    original = result
    corrected = result
    
    # 只對明顯的錯誤進行修正，避免過度修正
    obvious_corrections = {
        # 明顯的數字混淆
        'O': '0',   # 字母O改為數字0
        'l': '1',   # 小寫l改為數字1
        # 避免其他可能誤判的修正
    }
    
    # 檢查結果的合理性
    if len(result) > 8:  # 如果太長，可能識別錯誤
        print(f"識別結果過長 ({len(result)} 字符): {result}")
        # 嘗試取前6個字符
        corrected = result[:6]
    elif len(result) < 4:  # 如果太短，可能識別不完整
        print(f"識別結果過短 ({len(result)} 字符): {result}")
        # 保持原樣，讓系統重新識別
        return result
    
    # 應用明顯的修正
    for wrong, correct in obvious_corrections.items():
        if wrong in corrected:
            corrected = corrected.replace(wrong, correct)
    
    # 檢查是否有重複字符（可能是識別錯誤）
    if len(set(corrected)) < len(corrected) * 0.5:  # 如果唯一字符太少
        print(f"檢測到可能的重複字符錯誤: {corrected}")
        # 不進行修正，返回原結果
        return result
    
    if corrected != original:
        print(f"輕微修正: {original} -> {corrected}")
    
    return corrected


# 高效的驗證碼獲取（簡化版本）
def get_captcha_efficiently(chrome, i):
    try:
        # 點擊刷新驗證碼按鈕
        element = WebDriverWait(chrome, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="queryForm"]/div[1]/div[2]/div[2]/div/div/etw-captcha/div/button[1]')))
        element.click()
        time.sleep(0.5)
        
        # 下載驗證碼圖片
        img = download_captcha_image(chrome, i)
        if img is None:
            # 備用截圖方式
            chrome.save_screenshot(f'./img/screenshot{i}.png')
            page_snap_obj = Image.open(f'./img/screenshot{i}.png')
            img = page_snap_obj.crop((left, top, right, bottom))
        
        # 使用最有效的單一處理策略
        if img.mode != 'L':
            img = img.convert('L')
        
        # 簡單但有效的處理
        width, height = img.size
        img = img.resize((width * 3, height * 3), Image.Resampling.LANCZOS)
        
        # 基本二值化
        img = convert_img(img, 140)
        
        # 基本OCR識別
        result = pytesseract.image_to_string(img, config='--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz')
        clean_result = re.sub(r'[^a-zA-Z0-9]', '', result)
        
        # 如果結果不是6位，最多重試2次
        retry_count = 0
        while len(clean_result) != 6 and retry_count < 2:
            retry_count += 1
            element.click()
            time.sleep(0.3)
            img = download_captcha_image(chrome, f"{i}_retry{retry_count}")
            if img:
                if img.mode != 'L':
                    img = img.convert('L')
                img = img.resize((width * 3, height * 3), Image.Resampling.LANCZOS)
                img = convert_img(img, 140)
                result = pytesseract.image_to_string(img, config='--psm 8 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz')
                clean_result = re.sub(r'[^a-zA-Z0-9]', '', result)
        
        return clean_result if len(clean_result) >= 4 else ''
        
    except Exception as e:
        print(f"高效驗證碼獲取失敗: {e}")
        return ''

# 處理驗證碼驗證結果
def handle_captcha_verification(chrome, taxNo, i):
    try:
        # 檢查是否有錯誤對話框（驗證碼錯誤）
        info = find_element_or_return_empty(chrome,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
        
        retry_count = 0
        while len(info) > 0 and retry_count < 3:  # 最多重試3次
            retry_count += 1
            chrome.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button').click()
            time.sleep(0.3)
            
            # 重新獲取驗證碼
            result = get_captcha_efficiently(chrome, f"{i}_verify{retry_count}")
            if not result:
                return False
            
            # 重新填入統編和驗證碼
            taxNoBox = chrome.find_element(By.ID,'ban')
            captcha = chrome.find_element(By.ID,'captchaText')
            taxNoBox.clear()
            taxNoBox.send_keys(taxNo)
            captcha.clear()
            captcha.send_keys(result)
            captcha.submit()
            
            # 再次檢查是否有錯誤對話框
            info = find_element_or_return_empty(chrome,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
        
        return len(info) == 0  # 沒有錯誤對話框表示成功
        
    except Exception as e:
        print(f"驗證碼驗證處理失敗: {e}")
        return False

# 提取負責人姓名
def extract_responsible_name(chrome):
    try:
        # 找到包含 "負責人姓名" 的 <div>
        div_elements = chrome.find_elements(By.XPATH, "//div[contains(text(), '負責人姓名')]")
        for div in div_elements:
            # 找到 "負責人姓名" 的下一个兄弟节点的文本
            responsible_name = div.find_element(By.XPATH, "./following-sibling::div").text
            return responsible_name.strip()
        return None
    except Exception as e:
        print(f"提取負責人姓名失敗: {e}")
        return None

# 更新資料庫記錄（使用現有連線）
def update_database_record(db_connection, responsible_name, taxNo):
    try:
        cursor = db_connection.cursor()
        cursor.execute("UPDATE CrawlerData SET d=? WHERE a=?", (responsible_name, taxNo))
        db_connection.commit()
        cursor.close()
    except Exception as e:
        print(f"更新資料庫記錄失敗: {e}")

def find_element_or_return_empty(chrome,xpath):
    try:
        element = WebDriverWait(chrome, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return element.text
    except TimeoutException:
        return ''  # 或者根據需要返回其他適當的值

# 下載驗證碼圖片
def download_captcha_image(chrome, i):
    try:
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
            print(f"無法找到驗證碼圖片元素")
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
                print(f"無法下載驗證碼圖片，狀態碼: {response.status_code}")
                return None
        
        # 保存原始圖片
        img.save(f"./img/captcha_original{i}.png")
        print(f"成功下載驗證碼圖片: captcha_original{i}.png")
        return img
        
    except Exception as e:
        print(f"下載驗證碼圖片時發生錯誤: {e}")
        return None

# 增強驗證碼圖片處理
def enhance_captcha_image(img, i):
    try:
        if img is None:
            return None
            
        # 轉換為RGB格式（如果需要）
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # 保存原始圖片供調試
        img.save(f"./img/captcha_step0_original{i}.png")
        
        # 針對驗證碼特點進行預處理
        # 先進行顏色標準化
        img = normalize_captcha_colors(img)
        img.save(f"./img/captcha_step1_normalized{i}.png")
        
        # 調整圖片大小（根據原始大小智能縮放）
        width, height = img.size
        target_height = 40  # Tesseract最佳識別高度
        scale_factor = target_height / height
        new_width = int(width * scale_factor)
        img = img.resize((new_width, target_height), Image.Resampling.LANCZOS)
        img.save(f"./img/captcha_step2_resized{i}.png")
        
        # 增強對比度和銳度
        from PIL import ImageEnhance, ImageFilter
        
        # 轉換為灰度
        img = img.convert('L')
        img.save(f"./img/captcha_step3_gray{i}.png")
        
        # 使用多閾值策略
        best_img = multi_threshold_binarization(img, i)
        img.save(f"./img/captcha_step4_binary{i}.png")
        
        # 字符分離處理
        img = character_isolation(best_img, i)
        img.save(f"./img/captcha_step5_isolated{i}.png")
        
        # 最終清理
        img = final_cleanup_for_tesseract(img)
        img.save(f"./img/captcha_step6_final{i}.png")
        
        return img
        
    except Exception as e:
        print(f"增強驗證碼圖片時發生錯誤: {e}")
        # 回退到簡單處理
        return simple_enhance_fallback(img, i)

# 處理圖形驗證辨識
def refresh_captcha(chrome,i):
    try:
        # 點擊刷新驗證碼按鈕
        element = WebDriverWait(chrome, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="queryForm"]/div[1]/div[2]/div[2]/div/div/etw-captcha/div/button[1]')))
        element.click()
        time.sleep(0.8)  # 增加等待時間確保驗證碼完全載入
        
        # 嘗試多次下載驗證碼
        img = None
        for attempt in range(3):
            img = download_captcha_image(chrome, f"{i}_attempt{attempt}")
            if img is not None:
                break
            time.sleep(0.3)
        
        if img is None:
            print("下載驗證碼失敗，使用截圖方式")
            # 備用方案：使用截圖方式
            chrome.save_screenshot(f'./img/screenshot{i}.png')
            page_snap_obj = Image.open(f'./img/screenshot{i}.png')
            img = page_snap_obj.crop((left, top, right, bottom))
        
        # 使用多種處理策略
        results = []
        
        # 策略1：完整優化處理
        try:
            enhanced_img = enhance_captcha_image(img, f"{i}_enhanced")
            if enhanced_img is not None:
                result1 = smart_tesseract_recognition(enhanced_img, i)
                if result1:
                    results.append(result1)
                    print(f"優化處理結果: {result1}")
        except Exception as e:
            print(f"優化處理失敗: {e}")
        
        # 策略2：簡單處理
        try:
            simple_img = simple_enhance_fallback(img, f"{i}_simple")
            if simple_img is not None:
                result2 = smart_tesseract_recognition(simple_img, i)
                if result2:
                    results.append(result2)
                    print(f"簡單處理結果: {result2}")
        except Exception as e:
            print(f"簡單處理失敗: {e}")
        
        # 策略3：原始圖片直接識別
        try:
            if img.mode != 'L':
                gray_img = img.convert('L')
            else:
                gray_img = img
            result3 = pytesseract.image_to_string(gray_img)
            clean_result3 = re.sub(r'[^a-zA-Z0-9]', '', result3)
            if clean_result3:
                results.append(clean_result3)
                print(f"原始處理結果: {clean_result3}")
        except Exception as e:
            print(f"原始處理失敗: {e}")
        
        # 選擇最佳結果
        if results:
            # 優先選擇6位的結果
            six_char_results = [r for r in results if len(r) == 6]
            if six_char_results:
                best_result = six_char_results[0]
            else:
                # 選擇長度最接近6的結果
                best_result = min(results, key=lambda x: abs(len(x) - 6))
            
            # 後處理
            final_result = post_process_tesseract_result(best_result)
            print(f"最終識別結果: {final_result}")
            return final_result
        
        print("所有識別策略都失敗")
        return ''
        
    except Exception as e:
        print(f"處理圖形驗證時發生異常: {e}")
        return ''

# 处理单个数据库文件的函数（優化版本）
def process_database(db_filename, i):
    chrome = None
    mydb = None
    
    try:
        # 使用優化的 Chrome 實例
        chrome = create_optimized_chrome()
        
        # 建立資料庫連線
        mydb = sqlite3.connect(db_filename, check_same_thread=False, timeout=30)
        mydb.execute("PRAGMA journal_mode=WAL")
        mydb.execute("PRAGMA synchronous=NORMAL")
        
        batch_size = 20  # 每次處理的記錄數量
        
        while True:
            cursor = mydb.cursor()
            cursor.execute(f"SELECT a FROM CrawlerData where d='' and length(a)=8 ORDER BY RANDOM() LIMIT {batch_size}")
            rows = cursor.fetchall()
            cursor.close()
            
            if not rows:
                logger.info(f"資料庫 {db_filename} 已無待處理資料")
                break
                
            logger.info(f"資料庫 {i} 處理 {len(rows)} 筆資料")
            
            # 處理每批資料
            for row in rows:
                try:
                    monitor.record_processed()
                    taxNo = str(row[0])
                    
                    if process_single_record(chrome, mydb, taxNo, i):
                        monitor.record_success()
                    else:
                        monitor.record_error()
                        
                except Exception as e:
                    logger.error(f"處理統編 {row[0]} 時發生錯誤: {e}")
                    monitor.record_error()
                    continue
                    
    except Exception as e:
        logger.error(f"資料庫處理發生錯誤: {e}")
    finally:
        cleanup_resources(mydb, chrome)

# 處理單筆記錄
def process_single_record(chrome, mydb, taxNo, i):
    """處理單筆記錄"""
    try:
        chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
        
        taxNoBox = WebDriverWait(chrome, 10).until(EC.presence_of_element_located((By.ID, "ban")))
        taxNoBox.clear()
        taxNoBox.send_keys(taxNo)
        
        # 檢查是否有錯誤訊息
        try:
            chrome.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/div/div/jhi-main/etw113w1-ban-query/form/div[1]/div[2]/div[1]/div/div/small')
            return False
        except:
            pass
        
        # 簡化的驗證碼處理
        result = get_captcha_efficiently(chrome, i)
        if not result:
            logger.warning(f"統編 {taxNo} 驗證碼識別失敗")
            return False
        
        # 填入統編和驗證碼並提交
        taxNoBox = chrome.find_element(By.ID, 'ban')
        captcha = chrome.find_element(By.ID, 'captchaText')
        taxNoBox.clear()
        taxNoBox.send_keys(taxNo)
        captcha.clear()
        captcha.send_keys(result)
        captcha.submit()
        
        # 檢查提交結果
        if handle_captcha_verification(chrome, taxNo, i):
            responsible_name = extract_responsible_name(chrome)
            if responsible_name:
                update_database_record(mydb, responsible_name, taxNo)
                logger.info(f"成功更新統編 {taxNo}: {responsible_name}")
                chrome.back()
                return True
        
        chrome.back()
        return False
        
    except Exception as e:
        logger.error(f"處理統編 {taxNo} 時發生錯誤: {e}")
        return False

# 清理資源
def cleanup_resources(mydb, chrome):
    """清理資源"""
    if mydb:
        try:
            mydb.close()
        except:
            pass
    if chrome:
        try:
            chrome.quit()
        except:
            pass

# 創建優化的 Chrome 瀏覽器實例
def create_optimized_chrome():
    """創建優化配置的 Chrome 瀏覽器實例"""
    
    # 設置環境變數來禁用 GCM
    gcm_disable_vars = {
        'CHROME_HEADLESS': '1',
        'CHROME_NO_SANDBOX': '1',
        'GOOGLE_API_KEY': '',
        'GOOGLE_DEFAULT_CLIENT_ID': '',
        'GOOGLE_DEFAULT_CLIENT_SECRET': '',
        'CHROME_DISABLE_GCM': '1',
        'CHROME_DISABLE_SYNC': '1',
        'CHROME_DISABLE_BACKGROUND_NETWORKING': '1',
        'CHROME_DISABLE_COMPONENT_UPDATE': '1'
    }
    
    # 設置環境變數
    for key, value in gcm_disable_vars.items():
        os.environ[key] = value
    
    options = webdriver.ChromeOptions()
    
    # 核心禁用選項 - 針對 GCM 和 MCS 錯誤
    core_options = [
        '--headless=new',
        '--no-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--disable-software-rasterizer',
        '--disable-background-networking',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-sync',
        '--disable-translate',
        '--disable-features=GCMChannelStatusRequest,PushMessaging,Notifications',
        '--disable-notifications',
        '--disable-push-messaging',
        '--disable-gcm-registration',
        '--disable-background-mode',
        '--disable-component-update',
        '--disable-domain-reliability',
        '--disable-background-downloads',
        '--disable-cloud-import',
        '--disable-sync-types',
        '--no-pings',
        '--disable-web-resources',
        '--disable-component-cloud-policy',
        '--disable-extensions',
        '--disable-logging',
        '--log-level=3',
        '--silent',
        '--disable-breakpad',
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
        '--disable-back-forward-cache',
        '--disable-features=VizDisplayCompositor,TranslateUI,BlinkGenPropertyTrees',
        '--disable-features=AutofillServerCommunication,CertificateTransparencyComponentUpdater',
        '--disable-features=UserMediaScreenCapturing,MediaRouter,PasswordsAccountStorage',
        '--disable-features=NetworkService,VizServiceBase,WebRTC'
    ]
    
    for option in core_options:
        options.add_argument(option)
    
    # 設置用戶代理
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    # 強化的 prefs 設置
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.push_messaging": 2,
        "profile.default_content_setting_values.geolocation": 2,
        "profile.default_content_setting_values.media_stream": 2,
        "profile.default_content_setting_values.automatic_downloads": 2,
        "profile.default_content_setting_values.mixed_script": 2,
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
    
    # 排除開關和禁用自動化檢測
    options.add_experimental_option("excludeSwitches", [
        "enable-logging",
        "enable-automation",
        "enable-blink-features"
    ])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    return webdriver.Chrome(options=options)


# 创建一个线程列表
threads = []

# 增強的線程管理
def start_threads():
    global threads
    # 先清空線程列表
    threads = []
    # 排除的數字列表
    exclude_list = []
    
    # 減少並發線程數量以提升穩定性
    max_threads = 15
    
    for i in range(1, max_threads + 1):
        if i in exclude_list:
            continue  # 跳過這些值
        db_filename = f"D:/workspace/CrawlerProject/var/db{i}.db"
        
        # 檢查資料庫檔案是否存在
        if not os.path.exists(db_filename):
            print(f"資料庫檔案不存在: {db_filename}")
            continue
            
        thread = threading.Thread(target=process_database, args=(db_filename, i))
        thread.daemon = True  # 設為守護線程
        threads.append(thread)
        thread.start()
        time.sleep(0.1)  # 避免同時啟動造成競爭

# 優化的停止函數
def stop_threads():
    global threads
    print("正在停止所有線程...")
    for thread in threads:
        if thread.is_alive():
            thread.join(timeout=30)  # 設定超時時間
    threads = []

# 優化的重啟機制
def restart_threads():
    print("開始重啟線程...")
    stop_threads()
    time.sleep(2)  # 給予清理時間
    print("重新啟動所有線程...")
    start_threads()

# 優化的定時重啟（增加錯誤處理）
def schedule_restart(interval_hours):
    interval_seconds = interval_hours * 3600
    while True:
        try:
            time.sleep(interval_seconds)
            restart_threads()
        except Exception as e:
            logger.error(f"定時重啟發生錯誤: {e}")
            time.sleep(60)  # 錯誤後等待1分鐘再重試

# 主程式執行
def main():
    """主程式入口"""
    try:
        # 初始化
        logger.info("正在初始化爬蟲系統...")
        initialize_directories()
        start_monitoring()
        
        # 啟動線程
        logger.info("正在啟動處理線程...")
        start_threads()
        
        # 啟動定時重啟（4小時）
        logger.info("正在啟動定時重啟機制...")
        restart_scheduler = threading.Thread(target=schedule_restart, args=(4,))
        restart_scheduler.daemon = True
        restart_scheduler.start()
        
        logger.info("爬蟲系統已成功啟動！")
        
        # 保持主線程運行
        try:
            while True:
                time.sleep(60)  # 每分鐘檢查一次
                if not any(t.is_alive() for t in threads):
                    logger.warning("所有處理線程已停止，重新啟動...")
                    start_threads()
        except KeyboardInterrupt:
            logger.info("收到中斷信號，正在關閉系統...")
            stop_threads()
            logger.info("系統已安全關閉")
            
    except Exception as e:
        logger.error(f"主程式執行錯誤: {e}")
        stop_threads()

if __name__ == "__main__":
    main()



    
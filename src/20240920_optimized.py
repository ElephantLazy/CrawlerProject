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

# 全域變數
left = 615
top = 600
right = 793
bottom = 681
threshold = 150

# 初始化函數
def initialize_directories():
    """創建必要的目錄"""
    if not os.path.exists('./img'):
        os.makedirs('./img')
        logger.info("已創建 img 目錄")

# 優化的圖像二值化
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

# 下載驗證碼圖片
def download_captcha_image(chrome, i):
    try:
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
            logger.warning("無法找到驗證碼圖片元素")
            return None
            
        img_src = captcha_img.get_attribute('src')
        
        if img_src.startswith('data:image'):
            header, encoded = img_src.split(',', 1)
            img_data = base64.b64decode(encoded)
            img = Image.open(io.BytesIO(img_data))
        else:
            response = requests.get(img_src, stream=True)
            if response.status_code == 200:
                img = Image.open(io.BytesIO(response.content))
            else:
                logger.error(f"無法下載驗證碼圖片，狀態碼: {response.status_code}")
                return None
        
        img.save(f"./img/captcha_original{i}.png")
        return img
        
    except Exception as e:
        logger.error(f"下載驗證碼圖片時發生錯誤: {e}")
        return None

# 高效的驗證碼獲取（簡化版本）
def get_captcha_efficiently(chrome, i):
    try:
        # 點擊刷新驗證碼按鈕
        element = WebDriverWait(chrome, PerformanceConfig.ELEMENT_WAIT_TIME).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="queryForm"]/div[1]/div[2]/div[2]/div/div/etw-captcha/div/button[1]'))
        )
        element.click()
        time.sleep(PerformanceConfig.CAPTCHA_REFRESH_WAIT)
        
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
        img = img.resize((width * PerformanceConfig.IMAGE_SCALE_FACTOR, height * PerformanceConfig.IMAGE_SCALE_FACTOR), Image.Resampling.LANCZOS)
        
        # 基本二值化
        img = convert_img(img, PerformanceConfig.BINARY_THRESHOLD)
        
        # 基本OCR識別
        result = pytesseract.image_to_string(img, config=PerformanceConfig.TESSERACT_CONFIG)
        clean_result = re.sub(r'[^a-zA-Z0-9]', '', result)
        
        # 如果結果不是6位，最多重試指定次數
        retry_count = 0
        while len(clean_result) != 6 and retry_count < PerformanceConfig.CAPTCHA_RETRY_MAX:
            retry_count += 1
            element.click()
            time.sleep(PerformanceConfig.CAPTCHA_REFRESH_WAIT)
            img = download_captcha_image(chrome, f"{i}_retry{retry_count}")
            if img:
                if img.mode != 'L':
                    img = img.convert('L')
                img = img.resize((width * PerformanceConfig.IMAGE_SCALE_FACTOR, height * PerformanceConfig.IMAGE_SCALE_FACTOR), Image.Resampling.LANCZOS)
                img = convert_img(img, PerformanceConfig.BINARY_THRESHOLD)
                result = pytesseract.image_to_string(img, config=PerformanceConfig.TESSERACT_CONFIG)
                clean_result = re.sub(r'[^a-zA-Z0-9]', '', result)
        
        return clean_result if len(clean_result) >= 4 else ''
        
    except Exception as e:
        logger.error(f"高效驗證碼獲取失敗: {e}")
        return ''

# 處理驗證碼驗證結果
def handle_captcha_verification(chrome, taxNo, i):
    try:
        info = find_element_or_return_empty(chrome, '/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
        
        retry_count = 0
        while len(info) > 0 and retry_count < 3:
            retry_count += 1
            chrome.find_element(By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button').click()
            time.sleep(0.3)
            
            result = get_captcha_efficiently(chrome, f"{i}_verify{retry_count}")
            if not result:
                return False
            
            taxNoBox = chrome.find_element(By.ID, 'ban')
            captcha = chrome.find_element(By.ID, 'captchaText')
            taxNoBox.clear()
            taxNoBox.send_keys(taxNo)
            captcha.clear()
            captcha.send_keys(result)
            captcha.submit()
            
            info = find_element_or_return_empty(chrome, '/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
        
        return len(info) == 0
        
    except Exception as e:
        logger.error(f"驗證碼驗證處理失敗: {e}")
        return False

# 提取負責人姓名
def extract_responsible_name(chrome):
    try:
        div_elements = chrome.find_elements(By.XPATH, "//div[contains(text(), '負責人姓名')]")
        for div in div_elements:
            responsible_name = div.find_element(By.XPATH, "./following-sibling::div").text
            return responsible_name.strip()
        return None
    except Exception as e:
        logger.error(f"提取負責人姓名失敗: {e}")
        return None

# 更新資料庫記錄（使用現有連線）
def update_database_record(db_connection, responsible_name, taxNo):
    try:
        cursor = db_connection.cursor()
        cursor.execute("UPDATE CrawlerData SET d=? WHERE a=?", (responsible_name, taxNo))
        db_connection.commit()
        cursor.close()
    except Exception as e:
        logger.error(f"更新資料庫記錄失敗: {e}")

def find_element_or_return_empty(chrome, xpath):
    try:
        element = WebDriverWait(chrome, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return element.text
    except TimeoutException:
        return ''

# 處理單筆記錄
def process_single_record(chrome, mydb, taxNo, i):
    """處理單筆記錄"""
    try:
        chrome.get("https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
        
        taxNoBox = WebDriverWait(chrome, PerformanceConfig.ELEMENT_WAIT_TIME).until(
            EC.presence_of_element_located((By.ID, "ban"))
        )
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

# 处理单个数据库文件的函数（優化版本）
def process_database(db_filename, i):
    # 使用配置化的瀏覽器選項
    options = webdriver.ChromeOptions()
    browser_options = [
        '--headless', '--no-sandbox', '--disable-dev-shm-usage',
        '--disable-gpu', '--disable-extensions', '--disable-logging',
        '--disable-web-security', '--disable-features=VizDisplayCompositor',
        '--disable-images'  # 不載入圖片以節省頻寬
    ]
    for option in browser_options:
        options.add_argument(option)
    
    chrome = None
    mydb = None
    
    try:
        chrome = webdriver.Chrome(options=options)
        
        # 建立資料庫連線
        mydb = sqlite3.connect(db_filename, check_same_thread=False, timeout=30)
        mydb.execute("PRAGMA journal_mode=WAL")
        mydb.execute("PRAGMA synchronous=NORMAL")
        mydb.execute("PRAGMA cache_size=10000")
        
        while True:
            cursor = mydb.cursor()
            cursor.execute(f"SELECT a FROM CrawlerData where d='' and length(a)=8 ORDER BY RANDOM() LIMIT {PerformanceConfig.BATCH_SIZE}")
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

# 线程管理
threads = []

def start_threads():
    global threads
    threads = []
    exclude_list = []
    max_threads = min(PerformanceConfig.MAX_THREADS, 15)
    
    for i in range(1, max_threads + 1):
        if i in exclude_list:
            continue
        
        db_filename = f"D:/workspace/CrawlerProject/var/db{i}.db"
        if not os.path.exists(db_filename):
            logger.warning(f"資料庫檔案不存在: {db_filename}")
            continue
            
        thread = threading.Thread(target=process_database, args=(db_filename, i))
        thread.daemon = True
        threads.append(thread)
        thread.start()
        time.sleep(PerformanceConfig.THREAD_START_DELAY)

def stop_threads():
    global threads
    logger.info("正在停止所有線程...")
    for thread in threads:
        if thread.is_alive():
            thread.join(timeout=30)
    threads = []

def restart_threads():
    logger.info("開始重啟線程...")
    stop_threads()
    time.sleep(2)
    logger.info("重新啟動所有線程...")
    start_threads()

def schedule_restart(interval_hours):
    interval_seconds = interval_hours * 3600
    while True:
        try:
            time.sleep(interval_seconds)
            restart_threads()
        except Exception as e:
            logger.error(f"定時重啟發生錯誤: {e}")
            time.sleep(60)

# 主程式執行
if __name__ == "__main__":
    # 初始化
    initialize_directories()
    start_monitoring()
    
    # 啟動線程
    start_threads()
    
    # 啟動定時重啟（4小時）
    restart_scheduler = threading.Thread(target=schedule_restart, args=(4,))
    restart_scheduler.daemon = True
    restart_scheduler.start()
    
    logger.info("爬蟲系統已啟動")

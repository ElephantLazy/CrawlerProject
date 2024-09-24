import re
import threading
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
from bs4 import BeautifulSoup
import sqlite3
from selenium.common.exceptions import TimeoutException
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

#全域變數
left =  615
top = 600
right =  793
bottom = 681
# 二值化閾值
threshold = 150

#影像二值化
def convert_img(img, threshold):
    img = img.convert("L")  # 處理灰白
    pixels = img.load()
    for x in range(img.width):
        for y in range(img.height):
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


def find_element_or_return_empty(chrome,xpath):
    try:
        element = WebDriverWait(chrome, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return element.text
    except TimeoutException:
        return ''  # 或者根據需要返回其他適當的值

# 處理圖形驗證辨識
def refresh_captcha(chrome,i):
    try:
        js = "refreshCaptcha();"
        element = WebDriverWait(chrome, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="queryForm"]/div[1]/div[2]/div[2]/div/div/etw-captcha/div/button[1]')))
        element.click()
        time.sleep(0.1)
        chrome.save_screenshot('./img/screenshot'+str(i)+'.png')
        page_snap_obj = Image.open('./img/screenshot'+str(i)+'.png')
        image_obj = page_snap_obj.crop((left, top, right, bottom))
        img = convert_img(image_obj, threshold)
        img.save("./img/captcha1"+str(i)+".png")
        img = clearImg(img)
        img.save("./img/captcha2"+str(i)+".png")
        time.sleep(0.1)
        result = pytesseract.image_to_string(img)
        cop = re.compile("[^a-z^A-Z^0-9]")  
        result = cop.sub('', result)  
        return result
    except Exception as e:
        print("處理圖形驗證時發生異常：", e)
        return ''

# 处理单个数据库文件的函数
def process_database(db_filename,i):
    for j in range(1,9999999999):
        # 讀取資料庫
        mydb = sqlite3.connect(db_filename)
        cursor = mydb.cursor()
        cursor.execute("SELECT b FROM CrawlerData where e='' and length(b)=8 ORDER BY RANDOM()")
        rows = cursor.fetchall()
        cursor.close()
        if rows.count==0:
            break
        # 創建一個選項物件

        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        chrome = webdriver.Chrome(options=options)  

        # 將選項添加到選項物件中，例如，如果你需要使用headless模式，可以這樣添加
    

        # chrome.execute_script(js)
        # time.sleep(0.3)

        #查詢所有統編
        for row in rows:
            try:
                
                chrome.get(
                "https://www.etax.nat.gov.tw/etwmain/etw113w1/ban/result")
                # js = "var q=document.documentElement.scrollTop=50"
                result=''
                info=''
                # time.sleep(0.3)
                # taxNoBox = chrome.find_element(By.ID,"ban") 
                # 等待最多10秒直到元素出現
                taxNoBox = WebDriverWait(chrome, 2).until(EC.presence_of_element_located((By.ID, "ban")))
                taxNo = str(row[0])
                taxNoBox.clear()
                taxNoBox.send_keys(taxNo)
                chrome.implicitly_wait(1)
                chrome.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/div/div/jhi-main/etw113w1-ban-query/form/div[1]/div[2]/div[1]/div/div/small')
                continue
            except Exception as e:
                pass
            
            chrome.save_screenshot('./img/screenshot'+str(i)+'.png')
            page_snap_obj = Image.open('./img/screenshot'+str(i)+'.png')
            image_obj = page_snap_obj.crop((left, top, right, bottom))
            
            img = convert_img(image_obj, threshold)
            img.save("./img/captcha1"+str(i)+".png")
            img = clearImg(img)
            img.save("./img/captcha2"+str(i)+".png")
            result = pytesseract.image_to_string(img)
            # 只保留英數
            cop = re.compile("[^a-z^A-Z^0-9]")  
            result = cop.sub('', result) 
            while True:
                result = refresh_captcha(chrome,i)
                if len(result) == 6:
                    break
            #通過一次查詢
            taxNoBox = chrome.find_element(By.ID,'ban')
            captcha = chrome.find_element(By.ID,'captchaText')
            taxNoBox.send_keys(taxNo)
            captcha.send_keys(result)
            captcha.submit()
            # info = WebDriverWait(chrome, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')))
            info= find_element_or_return_empty(chrome,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
            #圖形驗證失敗
            while(len(info)>0):
                result=""
                chrome.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button').click()
                while True:
                    result = refresh_captcha(chrome,i)
                    if len(result) == 6:
                        break
                #通過一次查詢
                taxNoBox = chrome.find_element(By.ID,'ban')
                captcha = chrome.find_element(By.ID,'captchaText')
                taxNoBox.send_keys('70775250')
                captcha.send_keys(result)
                captcha.submit()
                # info = chrome.find_elements(By.XPATH,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
                info= find_element_or_return_empty(chrome,'/html/body/ngb-modal-window/div/div/jhi-dialog/div/div[3]/div/button')
            # 找到包含 "負責人姓名" 的 <div>
            div_elements = chrome.find_elements(By.XPATH, "//div[contains(text(), '負責人姓名')]")
            # 遍历找到的元素
            for div in div_elements:
                # 找到 "負責人姓名" 的下一个兄弟节点的文本
                responsible_name = div.find_element(By.XPATH, "./following-sibling::div").text
                mydb = sqlite3.connect(db_filename)
                cursor = mydb.cursor()
                cursor.execute("UPDATE CrawlerData SET e=? WHERE b=?", (responsible_name, taxNo))
                mydb.commit()
                cursor.close()
                # 使用CDP執行清理暫存的命令
                # chrome.execute_cdp_cmd("Network.clearBrowserCache", {})
                chrome.back()
                # chrome.close()
                # chrome = webdriver.Chrome(options=options)


# 创建一个线程列表
threads = []

# 啟動執行緒
def start_threads():
    global threads
    # 先清空線程列表
    threads = []
    # 排除的數字列表
    exclude_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
    for i in range(1, 16):
        if i in exclude_list:
            continue  # 跳過這些值
        db_filename = f"D:/CrawlerProject-main/var/db{i}.db"
        thread = threading.Thread(target=process_database, args=(db_filename, i))
        threads.append(thread)
        thread.start()

# 停止所有執行緒
def stop_threads():
    global threads
    for thread in threads:
        if thread.is_alive():
            print("Waiting for thread to finish...")
        thread.join()

# 重啟所有執行緒
def restart_threads():
    stop_threads()
    print("Restarting all threads...")
    start_threads()

# 每12小時重啟
def schedule_restart(interval_hours):
    interval_seconds = interval_hours * 3600
    while True:
        time.sleep(interval_seconds)
        restart_threads()

# 初始化第一次執行
start_threads()

# 啟動定時器每4小時重啟一次
restart_scheduler = threading.Thread(target=schedule_restart, args=(4,))
restart_scheduler.start()



    
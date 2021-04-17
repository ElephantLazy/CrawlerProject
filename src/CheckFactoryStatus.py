import re
# 圖片處理
from PIL import Image
# 文字識別
import pytesseract
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
import re
from bs4 import BeautifulSoup
import sqlite3
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
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

# 讀取資料庫
mydb = sqlite3.connect("./var/Factory.db")
cursor = mydb.cursor()
cursor2 = mydb.cursor()

options = Options()
chrome = webdriver.Chrome('./src/chromedriver', chrome_options=options)
chrome.get(
    "https://www.etax.nat.gov.tw/cbes/web/CBES113W1_1?vatId&fbclid=IwAR1ddyY-ipf3ZCnGLgxJijM62uyTLkIBqSaJeYHa6figtgztObIEuI-gElQ")
chrome.implicitly_wait(60)
img = chrome.find_element_by_xpath(
    '//*[@id="captchaImg"]')
js = "var q=document.documentElement.scrollTop=50"
# chrome.execute_script(js)
time.sleep(1)
location = img.location
size = img.size
left = location['x']+130
top = location['y']+125
right = left + size['width']+30
bottom = top + size['height']+5
chrome.save_screenshot('./img/screenshot.png')
page_snap_obj = Image.open('./img/screenshot.png')
image_obj = page_snap_obj.crop((left, top, right, bottom))
# 二值化閾值
threshold = 150
img = convert_img(image_obj, threshold)
img.save("./img/captcha1.png")
img = clearImg(img)
img.save("./img/captcha2.png")
result = pytesseract.image_to_string(img)
# 只保留英數
cop = re.compile("[^a-z^A-Z^0-9]")  
result = cop.sub('', result) 
while(len(result) != 6):
    js = "refreshCaptcha();"
    chrome.execute_script(js)
    chrome.save_screenshot('./img/screenshot.png')
    page_snap_obj = Image.open('./img/screenshot.png')
    image_obj = page_snap_obj.crop((left, top, right, bottom))
    img = convert_img(image_obj, threshold)
    img.save("./img/captcha1.png")
    img = clearImg(img)
    img.save("./img/captcha2.png")
    result = pytesseract.image_to_string(img)
    cop = re.compile("[^a-z^A-Z^0-9]")  # 匹配不是中文、大小写、数字的其他字符
    result = cop.sub('', result)  # 将string1中匹配到的字符替换成空字符
#通過一次查詢
taxNoBox = chrome.find_element_by_id("vatId")
captcha = chrome.find_element_by_id("captcha")
taxNoBox.send_keys('70775250')
captcha.send_keys(result)
captcha.submit()
info = chrome.find_elements_by_class_name("alert-danger")
#圖形驗證失敗
while(len(info)>3):
    result=""
    while(len(result) != 6):
        js = "refreshCaptcha();"
        chrome.execute_script(js)
        chrome.save_screenshot('./img/screenshot.png')
        page_snap_obj = Image.open('./img/screenshot.png')
        image_obj = page_snap_obj.crop((left, top, right, bottom))
        img = convert_img(image_obj, threshold)
        img.save("./img/captcha1.png")
        img = clearImg(img)
        img.save("./img/captcha2.png")
        result = pytesseract.image_to_string(img)
        cop = re.compile("[^a-z^A-Z^0-9]")  # 匹配不是中文、大小写、数字的其他字符
        result = cop.sub('', result)  # 将string1中匹配到的字符替换成空字符
    #通過一次查詢
    taxNoBox = chrome.find_element_by_id("vatId")
    captcha = chrome.find_element_by_id("captcha")
    taxNoBox.send_keys('70775250')
    captcha.send_keys(result)
    captcha.submit()
    info = chrome.find_elements_by_class_name("alert-danger")

#查詢所有統編
cursor2.execute("SELECT FactoryTaxNo FROM FactoryData WHERE IsSearch=0 and NeedSearch=1")
for row in cursor2:
    time.sleep(0.5)
    chrome.back()
    taxNoBox = chrome.find_element_by_id("vatId")
    taxNo = str(row[0])
    taxNoBox.clear()
    taxNoBox.send_keys(taxNo)
    taxNoBox.submit()
    info = chrome.find_elements_by_class_name("alert-danger")
    if(len(info)<3):
        cursor.execute("Update FactoryData set Status='"+info[0].text+"',IsSearch='1' WHERE FactoryTaxNo='"+taxNo+"'")
        mydb.commit()
    else:
        cursor.execute("Update FactoryData set Status='未知',IsSearch='1' WHERE FactoryTaxNo='"+taxNo+"'")
        mydb.commit()
        #圖形驗證失敗
        while(len(info)>3):
            result=""
            while(len(result) != 6):
                js = "refreshCaptcha();"
                chrome.execute_script(js)
                chrome.save_screenshot('./img/screenshot.png')
                page_snap_obj = Image.open('./img/screenshot.png')
                image_obj = page_snap_obj.crop((left, top, right, bottom))
                img = convert_img(image_obj, threshold)
                img.save("./captcha1.png")
                img = clearImg(img)
                img.save("./captcha2.png")
                result = pytesseract.image_to_string(img)
                cop = re.compile("[^a-z^A-Z^0-9]")  # 匹配不是中文、大小写、数字的其他字符
                result = cop.sub('', result)  # 将string1中匹配到的字符替换成空字符
            #通過一次查詢
            taxNoBox = chrome.find_element_by_id("vatId")
            captcha = chrome.find_element_by_id("captcha")
            taxNoBox.send_keys('70775250')
            captcha.send_keys(result)
            captcha.submit()
            info = chrome.find_elements_by_class_name("alert-danger")
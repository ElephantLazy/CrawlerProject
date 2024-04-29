import requests
import sqlite3
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
# 定义一个函数，用于处理单个数据库文件
def process_database(db_filename):
    mydb = sqlite3.connect(db_filename)
    cursor = mydb.cursor()
    cursor.execute("SELECT d FROM FactoryData where c=''")
    rows = cursor.fetchall()
    cursor.close()
    # 搜尋網站
    # 創建一個選項物件
    options = webdriver.ChromeOptions()

    # 將選項添加到選項物件中，例如，如果你需要使用headless模式，可以這樣添加
    chrome = webdriver.Chrome(options=options)
    chrome.get(
            "https://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do")
    chrome.implicitly_wait(1)
    chrome.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[3]').click()    
    chrome.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[5]').click()
    chrome.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[7]').click()
    chrome.find_element(By.XPATH,'/html/body/div[2]/form/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[9]').click()
    for row in rows:
        Business_Accounting_NO = row[0]
        Responsible_Name = ''
        try:
             # find iframe
            captcha_iframe = WebDriverWait(chrome, 10).until(
                ec.presence_of_element_located(
                    (
                        By.TAG_NAME, 'iframe'
                    )
                )
            )

            ActionChains(chrome).move_to_element(captcha_iframe).click().perform()

            # click im not robot
            captcha_box = WebDriverWait(chrome, 10).until(
                ec.presence_of_element_located(
                    (
                        By.ID, 'g-recaptcha-response'
                    )
                )
            )

            chrome.execute_script("arguments[0].click()", captcha_box)
            # 切回原始页面
            chrome.switch_to.default_content()
        except Exception as e:
            print(f"发生异常：{e}")
            # 切回原始页面
            chrome.switch_to.default_content()
        
        try:
            chrome.find_element(By.XPATH,'/html/body/table/tbody/tr/td/table/tbody/tr/td/span/a').click()
        except Exception as e:
            print(e)
            
        

        
        taxNoBox = chrome.find_element(By.ID,'qryCond')
        taxNoBox.send_keys(row[0])    
        js = "sendQueryList()"
        time.sleep(2)
        chrome.execute_script(js)
        
        try:
            detailBtn=chrome.find_element(By.XPATH,'//*[@id="vParagraph"]/div/div[2]/span[9]')
            time.sleep(2)
            detailBtn.click()
            td_elements=chrome.find_elements(By.TAG_NAME,'td')
            # 使用循环查找代表人姓名的下一个元素
            for index, td in enumerate(td_elements):
                if td.text == '代表人姓名':
                    # 找到了代表人姓名的td元素，获取下一个td元素的文本
                    Responsible_Name = td_elements[index + 1].text
                    break  # 找到后退出循环
            mydb = sqlite3.connect(db_filename)
            cursor = mydb.cursor()
            cursor.execute("UPDATE FactoryData SET c=? WHERE d=?", (Responsible_Name, Business_Accounting_NO))
            mydb.commit()
            cursor.close()
            chrome.back()
        except NoSuchElementException:
            print(f"找不到元素，跳过处理：{Business_Accounting_NO}")
            continue
        
    mydb.close()

# 创建一个线程列表
threads = []

# 启动多个线程，每个线程处理一个数据库文件
for i in range(1, 2):
    db_filename = f"./var/db{i}.db"
    thread = threading.Thread(target=process_database, args=(db_filename,))
    threads.append(thread)
    thread.start()

# 等待所有线程完成
for thread in threads:
    thread.join()

print("所有数据库处理完成")


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import sqlite3

def CheckTaxNo():
    while True:
        try:
            # 讀取資料庫
            mydb = sqlite3.connect("../var/Factory.db")
            cursor = mydb.cursor()
            cursor.execute("SELECT FactoryTaxNo FROM FactoryData WHERE NeedSearch=0")
            cursor2 = mydb.cursor()
            options = Options()
            #options.add_argument('--headless') 
            # options.add_argument("--disable-notifications")
            chrome = webdriver.Chrome('./chromedriver', chrome_options=options)
            chrome.get(
                "https://www.etax.nat.gov.tw/cbes/web/CBES113W1_1?vatId&fbclid=IwAR1ddyY-ipf3ZCnGLgxJijM62uyTLkIBqSaJeYHa6figtgztObIEuI-gElQ")
            chrome.implicitly_wait(60)
            for row in cursor:
                taxNo = str(row[0])
                taxNoBox = chrome.find_element_by_id("vatId")
                taxNoBox.clear()
                taxNoBox.send_keys(taxNo)
                taxNoBox.submit()
                info = chrome.find_elements_by_class_name("alert-danger")
                if(info[0].text != '請輸入正確的資料'):
                    cursor2.execute(
                        "Update FactoryData set NeedSearch=1 WHERE FactoryTaxNo='"+taxNo+"'")
                    mydb.commit()
        except:
            if('chrome' in globals()):
                chrome.quit()

if __name__ == '__main__':
    CheckTaxNo()


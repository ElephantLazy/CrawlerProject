from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
import openpyxl
import sqlite3
from retrying import retry
import concurrent.futures
import random
from fake_useragent import UserAgent
import requests
from requests.adapters import HTTPAdapter
@retry
def crawlerCompany(index):
    while True:
        try:
            s = requests.Session()
            s.mount('http://', HTTPAdapter(max_retries=5))
            s.mount('https://', HTTPAdapter(max_retries=5))
            startRandom = random.randint(1, 10)
            searchRandom = random.randint(2, 5)
            # 讀取資料庫
            mydb = sqlite3.connect("../var/Company"+str(index)+".db")
            cursor = mydb.cursor()
            cursor.execute("SELECT value FROM GCIS_Company WHERE key='pageIndex'")
            for row in cursor:
                pageIndex = int(row[0])
            cursor.execute("SELECT value FROM GCIS_Company WHERE key='infoPageIndex'")
            for row in cursor:
                infoPageIndex = int(row[0])
            cursor.execute("SELECT value FROM GCIS_Company WHERE key='cityIndex'")
            for row in cursor:
                cityIndex = int(row[0])
            cursor.execute(
                "SELECT value FROM GCIS_Company WHERE key='mainSelectIndex'")
            for row in cursor:
                mainSelectIndex = int(row[0])
            cursor.execute("SELECT value FROM GCIS_Company WHERE key='selectIndex'")
            for row in cursor:
                selectIndex = int(row[0])

            options = Options()
            # options.add_argument('--headless') 
            # options.add_argument('--disable-gpu')
            # time.sleep(startRandom)
            chrome = webdriver.Chrome('./chromedriver', chrome_options=options)
            chrome.get(
                "https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do;jsessionid=6AC3B69556A77E80B2D953B9C94F6B81")
            chrome.implicitly_wait(60)
            # 獲取登記現況選項
            statusBoxs = chrome.find_elements_by_name("isAlive")
            time.sleep(2)
            statusBoxs[0].click()
            js = "javascript:toggleAdvSearch('turnOn')"
            chrome.execute_script(js)
            # 填入搜尋欄
            searchBox = chrome.find_element_by_id('qryCond')
            searchBox.send_keys('公司')
            # 取得鎖營事業
            mainSelect = Select(chrome.find_element_by_id('busiItemMain'))
            mainSelectCount = len(mainSelect.options)
            isFirstCycle = True
            for cityIndex in range(cityIndex, 22, 1):
                cityBoxs = chrome.find_elements_by_name("city")
                time.sleep(searchRandom)
                cityBoxs[cityIndex].click()
                if(isFirstCycle == False):
                    cityBoxs[cityIndex-5].click()
                isFirstCycle = False
                # 走訪所營事業
                for mainSelectIndex in range(mainSelectIndex, len(mainSelect.options), 1):
                    mainSelect = Select(chrome.find_element_by_id('busiItemMain'))
                    time.sleep(searchRandom)
                    mainSelect.select_by_index(mainSelectIndex)
                    select = Select(chrome.find_element_by_id('busiItemSub'))
                    selectCount = len(select.options)
                    # 走訪所營事業細項
                    for selectIndex in range(selectIndex, len(select.options), 1):
                        select = Select(chrome.find_element_by_id('busiItemSub'))
                        time.sleep(searchRandom)
                        select.select_by_index(selectIndex)
                        # 送出搜尋
                        js = "sendQueryList()"
                        time.sleep(searchRandom)
                        chrome.execute_script(js)
                        # 取得分頁數量
                        pageSoup = BeautifulSoup(chrome.page_source, 'html.parser')
                        pageInfo = pageSoup.find_all('div', {
                            'style': 'float:left'})        

                        if(len(pageInfo) > 0):
                            startIndex = pageInfo[0].text.find('分')
                            endIndex = pageInfo[0].text.find('頁')
                            pageCount = int(pageInfo[0].text[startIndex+1:endIndex])
                            for i in range(pageIndex, pageCount+1, 1):
                                if(pageCount > 1):
                                    js = "gotoPage("+str(i)+");"
                                    time.sleep(2)
                                    chrome.execute_script(js)
                                # 取得單頁所有公司詳細資訊
                                factorySoup = BeautifulSoup(
                                    chrome.page_source, 'html.parser')
                                factoryInfos=factorySoup.find_all(
                                    'div', {'class': 'panel panel-default'})
                                currentInfoIndex = 0
                                oldTaxNo=""
                                for f in factoryInfos:                                   
                                    if(currentInfoIndex >= infoPageIndex-1):
                                        #取得統編
                                        startIndex = f.text.find('統一編號：')
                                        searchTaxNo=""
                                        searchTaxNo=f.text[startIndex+5:startIndex+13]
                                        newTaxNo=searchTaxNo
                                        if(newTaxNo==oldTaxNo):
                                            print("重複資料")
                                        oldTaxNo=searchTaxNo
                                        # 取得公司資訊
                                        time.sleep(0.3)
                                        r = s.get('http://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq '+searchTaxNo+'&$skip=0&$top=50', timeout=3)
                                        if(len(r.text)>0):
                                            companyTaxNo=searchTaxNo
                                            companyName=r.json()[0]['Company_Name']
                                            companyAddress=r.json()[0]['Company_Location'] 
                                            companyPerson=r.json()[0]['Responsible_Name']                                            
                                            companyCapital=str(format(r.json()[0]['Capital_Stock_Amount'],","))                                        
                                            companyStatus=r.json()[0]['Company_Status_Desc']                                            
                                            paperYear=r.json()[0]['Change_Of_Approval_Data']                                                                                               
                                            #Write to DB                                       
                                            cursor.execute(
                                                "INSERT INTO CompanyData VALUES ('"+companyName+"','"+companyTaxNo+"','"+companyStatus+"','"+companyCapital+"','"+companyPerson+"','"+companyAddress+"','"+paperYear+"','');")                                   
                                            mydb.commit()
                                        infoPageIndex += 1
                                        # 更新infoPageIndex
                                        cursor.execute("UPDATE GCIS_Company set value =" +
                                                    str(infoPageIndex)+" where key='infoPageIndex'")
                                        mydb.commit()
                                    currentInfoIndex += 1
                                # 更新infoPageIndex
                                infoPageIndex = 1
                                cursor.execute("UPDATE GCIS_Company set value =" +
                                            str(infoPageIndex)+" where key='infoPageIndex'")
                                mydb.commit()
                                cursor.execute("UPDATE GCIS_Company set value =" +
                                            str(pageIndex+1)+" where key='pageIndex'")
                                mydb.commit()
                                pageIndex += 1

                        # 重置pageIndex
                        pageIndex = 1
                        cursor.execute(
                            "UPDATE GCIS_Company set value ='1' where key='pageIndex'")
                        mydb.commit()
                        cursor.execute("UPDATE GCIS_Company set value =" +
                                    str(selectIndex+1)+" where key='selectIndex'")
                        mydb.commit()
                    # 重置selectIndex
                    cursor.execute(
                        "UPDATE GCIS_Company set value ='1' where key='selectIndex'")
                    mydb.commit()
                    selectIndex = 1
                    cursor.execute("UPDATE GCIS_Company set value =" +
                                str(mainSelectIndex+1)+" where key='mainSelectIndex'")
                    mydb.commit()

                # 重置mainSelectIndex
                cursor.execute(
                    "UPDATE GCIS_Company set value ='1' where key='mainSelectIndex'")
                mydb.commit()
                mainSelectIndex = 1
                # 更新cityIndex
                cursor.execute("UPDATE GCIS_Company set value =" +
                            str(cityIndex+1)+" where key='cityIndex'")
                mydb.commit()
        except Exception as e:
            print(e)
        finally:
            time.sleep(5)
            chrome.quit()
    


if __name__ == '__main__':
    crawlerCompany(0)
    # indexs = [0,1,2,3,4]
    # # 同時建立及啟用5個執行緒
    # with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
    #     executor.map(crawlerCompany, indexs)

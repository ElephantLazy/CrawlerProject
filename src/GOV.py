from cmath import log
from datetime import datetime
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
import requests
from requests.adapters import HTTPAdapter
import calendar
@retry
def crawlerCompany(index):
    while True:
        try:
            delay_choices = [3, 5, 2, 1, 4, 6]
            s = requests.Session()
            s.mount('http://', HTTPAdapter(max_retries=5))
            s.mount('https://', HTTPAdapter(max_retries=5))
            startRandom = random.randint(1, 10)
            searchRandom = random.randint(2, 5)
            # 讀取資料庫
            mydb = sqlite3.connect("./var/GOV.db")
            cursor = mydb.cursor()
            cursor.execute("SELECT value FROM Variable WHERE key='pageIndex'")
            for row in cursor:
                pageIndex = int(row[0])
            cursor.execute("SELECT value FROM Variable WHERE key='infoIndex'")
            for row in cursor:
                infoIndex = int(row[0])    
            cursor.execute("SELECT value FROM Variable WHERE key='radioIndex'")
            for row in cursor:
                radioIndex = int(row[0])    
            cursor.execute("SELECT value FROM Variable WHERE key='dateIndex'")
            for row in cursor:
                dateIndex = row[0]
            dateList=['2020/01/01','2020/02/01','2020/03/01','2020/04/01',
            '2020/05/01','2020/06/01','2020/07/01','2020/08/01',
            '2020/09/01','2020/10/01','2020/11/01','2020/12/01',
            '2021/01/01','2021/02/01','2021/03/01','2021/04/01',
            '2021/05/01','2021/06/01','2021/07/01','2021/08/01',
            '2021/09/01','2021/10/01','2021/11/01','2021/12/01',
            '2022/01/01','2022/02/01','2022/03/01']
          
            options = Options()
            options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36')

            # options.add_argument('--headless') 
            # options.add_argument('--disable-gpu')
            # time.sleep(startRandom)
            # chrome = webdriver.Chrome('./src/chromedriver', options=options)
            chrome = webdriver.Chrome('./src/chromedriver', options=options)
            with open('./src/stealth.min.js') as f:
                js = f.read()
            chrome.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": js})
            chrome.get(
                "http://web.pcc.gov.tw/prkms/prms-searchProctrgClient.do?root=tps&fbclid=IwAR2av1VCCci-suNX6DF1-j-wRFEGruklnV1wHv55dDujTjwdsGGtoIp3otU")
            chrome.implicitly_wait(60)
            for radioIndex in range(radioIndex, 2, 1):
                # 獲取標的分類選項
                radioList = chrome.find_elements_by_name("pmsProctrgCate")
                radioList[radioIndex].click()
                for dateIndex in range(dateIndex, 26, 1):
                    startDate= datetime.strptime(dateList[dateIndex], '%Y/%m/%d')
                    lastDayOfMonth=calendar.monthrange(startDate.year,startDate.month)[1]
                    endDate=datetime.strptime(str(startDate.year)+'/'+str(startDate.month)+'/'+str(lastDayOfMonth), '%Y/%m/%d')
                    # 公告日期CheckBox選擇日期
                    checkTenderCustom=chrome.find_element_by_id('checkTenderCustom')
                    checkTenderCustom.click()
                    # 填入起始日期
                    tenderStartDate = chrome.find_element_by_id('tenderStartDate')
                    tenderStartDate.clear()
                    tenderStartDate.send_keys(str(startDate.year-1911)+'/'+str(startDate.month).zfill(2)+
                    '/'+str(startDate.day).zfill(2))
                    tenderEndDate = chrome.find_element_by_id('tenderEndDate')
                    tenderEndDate.clear()
                    tenderEndDate.send_keys(str(endDate.year-1911)+'/'+str(endDate.month).zfill(2)+
                    '/'+str(endDate.day).zfill(2))
                    searchButton=chrome.find_element_by_id('search')
                    searchButton.click()
                    # 取得頁數
                    pageSoup = BeautifulSoup(chrome.page_source, 'html.parser')
                    totalPageIndex = int(int(pageSoup.find_all('span', {'class': 'T11b'})[0].text)/100+1)
                    for pageIndex in range(pageIndex, totalPageIndex, 1):
                        # 走訪頁數
                        # 取得單頁所有詳細資訊
                        detailList=chrome.find_elements_by_tag_name('a')
                        for infoIndex in range(infoIndex, len(detailList), 2):
                            detailList=chrome.find_elements_by_tag_name('a')
                            detailList[101+infoIndex].click()
                            companySoup = BeautifulSoup(chrome.page_source, 'html.parser')
                            companyInfos = companySoup.find_all(
                                        'td', {'class': 'newstop'})
                            AgencyCode=companyInfos[0].text.strip()
                            AgencyName=companyInfos[1].text.strip()
                            CompanyName=companyInfos[2].text.strip()
                            AgencyAddress=companyInfos[3].text.strip()
                            ContactName=companyInfos[4].text.strip()
                            Tel=companyInfos[5].text.strip()
                            Tax=companyInfos[6].text.strip()
                            Mail=companyInfos[7].text.strip()
                            TenderName=companyInfos[9].text.strip()
                            Category=companyInfos[10].text.strip()
                            #Write to DB
                            cursor.execute(
                                "INSERT INTO CompanyData VALUES ('"+AgencyCode+"','"+AgencyName+"','"+CompanyName+"','"+AgencyAddress+
                                "','"+ContactName+"','"+Tel+"','"+Tax+"','"+Mail+"','"+TenderName+"','"+Category+"');")
                            mydb.commit()

                            # 更新infoIndex
                            cursor.execute("UPDATE Variable set value ='" +
                                        str(infoIndex+2)+"' where key='infoIndex'")
                            mydb.commit()
                            if(pageIndex==1):
                                chrome.get('http://web.pcc.gov.tw/tps/pss/tender.do?searchMode=common&searchType=basic&method=search')
                            else:
                                chrome.get('http://web.pcc.gov.tw/tps/pss/tender.do?searchMode=common&searchType=basic&method=search&isSpdt=&pageIndex='+pageIndex)


                    # 更新dateIndex
                    cursor.execute("UPDATE Variable set value ='" +
                                str(dateIndex+1)+"' where key='dateIndex'")
                    mydb.commit()
                    
                    # 重置pageIndex
                    cursor.execute("UPDATE Variable set value ='0' where key='pageIndex'")
                    mydb.commit()
                    # 重置infoIndex
                    cursor.execute("UPDATE Variable set value ='0' where key='infoIndex'")
                    mydb.commit()
                
                # 更新radioIndex
                cursor.execute("UPDATE Variable set value ='" +
                            str(radioIndex+1)+"' where key='radioIndex'")
                mydb.commit()
                # 重置dateIndex
                cursor.execute("UPDATE Variable set value ='0' where key='dateIndex'")
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

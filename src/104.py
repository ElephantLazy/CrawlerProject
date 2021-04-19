from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
import openpyxl
import sqlite3
import math
from selenium.webdriver.common.action_chains import ActionChains
def Crwler104():
    while True:
        try:
            # 讀取資料庫
            mydb = sqlite3.connect("../var/104.db")
            cursor = mydb.cursor()
            cursor.execute("SELECT value FROM CrawlerInfos WHERE key='industryIndex'")
            for row in cursor:
                industryIndex = int(row[0])

            cursor.execute("SELECT value FROM CrawlerInfos WHERE key='companyPageIndex'")
            for row in cursor:
                companyPageIndex = int(row[0])

            cursor.execute(
                "SELECT value FROM CrawlerInfos WHERE key='companyPageInfoIndex'")
            for row in cursor:
                companyPageInfoIndex = int(row[0])

            options = Options()
            options.add_argument('--headless') 
            chrome = webdriver.Chrome('./chromedriver', chrome_options=options)
            chrome.get("http://www.104.com.tw/jb/104i/custlist/list?order=8")
            chrome.implicitly_wait(60)
            # 將產業列表展開
            industryButton = chrome.find_element_by_xpath(
                "/html/body/div[3]/div[2]/div/div[1]/div[1]/h4")
            industryButton.click()
            for ind in range(industryIndex, 49, 1):
                # 獲取產業列表
                industryLinks = chrome.find_elements_by_class_name("a2")
                chrome.get(industryLinks[industryIndex].get_attribute('href'))
                # 取得總頁數
                resultCnt = int(chrome.find_element_by_class_name(
                    'resultCnt').text)
                totalPage = math.ceil(resultCnt/15)
                # 取得當下網址
                currentUrl = str(chrome.current_url)
                # 跳至紀錄頁數
                if(companyPageIndex!=1):
                    recordUrl = currentUrl+"&page="+str(companyPageIndex)
                    chrome.get(recordUrl)
                # 走訪產業下所有頁數
                for cp in range(companyPageIndex, totalPage, 1):     
                    # 獲取此頁所有公司資連結
                    time.sleep(1)
                    companys = chrome.find_elements_by_class_name("a5")
                    # 獲取此頁所有工作機會連結
                    allJobSoup = BeautifulSoup(
                        chrome.page_source, 'html.parser')
                    allJobList = allJobSoup.find_all(
                        'a', {'class': 'a2', 'target': '_blank'})
                    chekHaveJobSoup= allJobSoup.find_all(
                        'div', {'class': 'm-box w-resultBox'})
                    if(len(allJobList)>0):
                        # 走訪此頁所有公司資訊
                        currentCompanyPageInfoIndex = 1
                        cIndex=0
                        for c in companys:                     
                            if(currentCompanyPageInfoIndex >= companyPageInfoIndex):
                                # 開啟公司資訊分頁
                                js = "window.open('"+c.get_attribute('href')+"')"
                                chrome.execute_script(js)
                                # 切換窗口
                                allHandles = chrome.window_handles
                                routeWindowHandle = 0
                                routeTitle = 'None'
                                oringalHandles = allHandles[0]
                                newHandles = allHandles[1]
                                chrome.switch_to_window(newHandles)
                                routeTitle = chrome.title
                                # 取得公司資訊
                                companySoup = BeautifulSoup(
                                    chrome.page_source, 'html.parser')
                                companyInfos = companySoup.find_all(
                                    'span', {'class': 'condition'})
                                companyName = chrome.find_element_by_xpath(
                                    "/html/body/div[3]/div[1]/div/h1").text
                                industry = companyInfos[0].text.strip()
                                employeeCount = companyInfos[1].text.strip()
                                captial = companyInfos[2].text.strip()
                                contactPerson = companyInfos[3].text.strip()
                                address = companyInfos[4].text.strip()
                                companyUrl=''
                                if(len(companyInfos)>7):
                                    companyUrl = companyInfos[7].text.strip()
                                # 查詢統編
                                taxNo=""
                                searchCompanyName=companyName.split("_")
                                for sc in searchCompanyName:
                                    taxNo=""
                                    cursor.execute(
                                        "SELECT taxNo FROM Company WHERE companyName Like '%"+sc+"%'")
                                    for row in cursor:
                                        taxNo = str(row[0])
                                    if(taxNo!=""):
                                        break

                            
                                if('目前暫無工作機會' not in chekHaveJobSoup[cIndex].text):
                                    # 連結至所有工作機會
                                    jobUrl = "http://www.104.com.tw"
                                    jobUrl += allJobList[cIndex]['href']
                                    chrome.get(jobUrl)
                                    # 撈取所有工作機會
                                    jobSoup = BeautifulSoup(
                                        chrome.page_source, 'html.parser')
                                    # 取得頁數資訊page
                                    page = jobSoup.find_all(
                                        'div', {'class': 'w-pager'})
                                    startIndex=page[0].text.find('下')
                                    totalPage='1'
                                    if(len(page[0].text)>14):
                                        totalPage=page[0].text[startIndex-2:startIndex]
                                    elif(len(page[0].text)>1):
                                        totalPage=page[0].text[startIndex-1:startIndex]
                                        
                                
                                    currentJobInfoIndex = 1           
                                    chrome.get(jobUrl)
                                    isFirstPage=True
                                    for p in range(1,int(totalPage)+1,1):  
                                        cboxClose=chrome.find_element_by_id('cboxClose')
                                        if(cboxClose.text=="close"):
                                            ac = ActionChains(chrome)
                                            ac.move_to_element(cboxClose).move_by_offset(x_offset, y_offset).click().perform()
                                        if(isFirstPage==False):
                                            time.sleep(1)
                                            nextPage=chrome.find_element_by_link_text('下一頁 »')
                                            nextPage.click()
                                        isFirstPage=False
                                        # 撈取所有工作機會
                                        jobSoup = BeautifulSoup(
                                        chrome.page_source, 'html.parser')
                                        jobInfos = jobSoup.find_all(
                                        'div', {'itemtype': 'http://schema.org/JobPosting'})
                                        titleInfos= jobSoup.find_all(
                                            'span', {'itemprop': 'title'})
                                        datePostedInfos= jobSoup.find_all(
                                            'span', {'itemprop': 'datePosted'})
                                        tIndex=0 
                                        
                                        for j in jobInfos:                    
                                            jobTitle = str(titleInfos[tIndex].text).replace('"', "")
                                            jobTitle.strip("'")
                                            jobDate = datePostedInfos[tIndex].text 
                                            tIndex+=1                   
                                            # 寫入DB
                                            cursor.execute("Insert into Jobs Values('"+companyName+"','"+taxNo+"','"+industry+"','"+contactPerson+"','"+captial+"','"+employeeCount+"','"+address+"','"+jobTitle+" "+jobDate+"','"+companyUrl+"')")
                                            mydb.commit()
                                        time.sleep(0.5)                                                      
                                    chrome.close()
                                    chrome.switch_to_window(oringalHandles)
                                cIndex+=1
                                companyPageInfoIndex += 1
                                # 更新companyPageInfoIndex
                                cursor.execute("UPDATE CrawlerInfos set value =" +
                                            str(companyPageInfoIndex)+" where key='companyPageInfoIndex'")
                                mydb.commit()
                            else:
                                if('目前暫無工作機會' not in chekHaveJobSoup[cIndex].text):
                                    cIndex+=1
                            currentCompanyPageInfoIndex += 1        
                    
                    #下一頁
                    companyPageNext = chrome.find_element_by_link_text('下一頁 »')
                    companyPageNext.click()
                    # 重置companyPageInfoIndex
                    companyPageInfoIndex=1
                    # 更新companyPageIndex
                    cursor.execute("UPDATE CrawlerInfos set value =" +
                                str(companyPageInfoIndex)+" where key='companyPageInfoIndex'")
                    mydb.commit()
                    # 更新companyPageIndex
                    cursor.execute("UPDATE CrawlerInfos set value =" +
                                str(companyPageIndex+1)+" where key='companyPageIndex'")
                    mydb.commit()

                # 更新industryIndex
                cursor.execute("UPDATE CrawlerInfos set value =" +
                            str(industryIndex+1)+" where key='industryIndex'")
                mydb.commit()
        except Exception as e:
            print(e)
        finally:
            chrome.quit()

if __name__ == '__main__':
    Crwler104()

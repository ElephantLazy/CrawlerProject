from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
import openpyxl
import sqlite3
import math
from selenium.webdriver.common.action_chains import ActionChains
import requests
from requests.adapters import HTTPAdapter
import re
def Crwler104():
    while True:
        try:
          
            mydb = sqlite3.connect("../var/1111.db")
            cursor = mydb.cursor()
            cursor.execute(
                "SELECT value FROM CrawlerInfos WHERE key='cityIndex'")
            for row in cursor:
                cityIndex = int(row[0])

            cursor.execute(
                "SELECT value FROM CrawlerInfos WHERE key='areaIndex'")
            for row in cursor:
                areaIndex = int(row[0])

            cursor.execute(
                "SELECT value FROM CrawlerInfos WHERE key='companyPageIndex'")
            for row in cursor:
                companyPageIndex = int(row[0])
            
            cursor.execute(
                "SELECT value FROM CrawlerInfos WHERE key='companyPageInfoIndex'")
            for row in cursor:
                companyPageInfoIndex = int(row[0])

            options = Options()
            # options.add_argument('--headless')
            chrome = webdriver.Chrome(
                './chromedriver', chrome_options=options)       
            chrome.get("https://www.1111.com.tw/search/corp")
            chrome.implicitly_wait(60)
            js = "Util1111.TriggerEvent(document.getElementById('_corp_coCht'), 'click')"
            chrome.execute_script(js)
        
            #將地區展開
            citys = chrome.find_elements_by_class_name(
                "tcode-col-2")
            #走訪地區
            for c in range(cityIndex,len(citys)-6,1):
                citys[c].click()
                time.sleep(2)
                areaSoup = BeautifulSoup(
                            chrome.page_source, 'html.parser')
                allAreas = areaSoup.find_all(
                'label', {'class': 'tcode-col-2 tcode-md-2 tcode__panel-item'})
                for i in range(areaIndex,len(allAreas),1):
                    url="https://www.1111.com.tw/search/corp?co="+allAreas[i]['k']
                    chrome.get(url)
                    pageInfo=chrome.find_element_by_xpath("/html/body/div[8]/div/div/div[2]/select").text.split()
                    totalPage=int(pageInfo[len(pageInfo)-1])
                    pageIndex=0
                    #走訪頁數
                    for t in range(companyPageIndex,totalPage,1):
                        #取得單頁公司資訊
                        s = requests.Session()
                        s.mount('http://', HTTPAdapter(max_retries=5))
                        s.mount('https://', HTTPAdapter(max_retries=5))
                        r = s.get('https://www.1111.com.tw/search/corp?co='+allAreas[i]['k']+'&page='+str(t)+'&act=load_page', timeout=3)
                        startIndex=r.text.find("html")+7
                        endIndex=len(r.text)-3
                        testStr=r.text[startIndex:endIndex]
                        testStr=testStr.replace("\\",'')
                        companyLinksSoup = BeautifulSoup(
                                    testStr, 'html.parser')
                        companyLinks = companyLinksSoup.find_all(
                        'a', {'class': 'item__corp-organ--link'})
                        #走訪公司資訊
                        for c in range(companyPageInfoIndex,len(companyLinks),1): 
                            #進入公司頁面                          
                            chrome.get(companyLinks[c]['href'])
                            companySoup = BeautifulSoup(
                                    chrome.page_source, 'html.parser')
                            #公司名稱
                            companyName=companySoup.find_all(
                                'h1')[0].text
                            #公司資訊標題
                            companyInfoTitles=companySoup.find_all(
                                'div', {'class': 'dt'})
                            #公司資訊內容
                            companyInfoDatas=companySoup.find_all(
                                'div', {'class': 'dd'})
                            industry=''
                            taxNo=''
                            person=''
                            capital=''
                            employeeCount=''
                            cop = re.compile("[^\u4e00-\u9fa5^a-z^A-Z^0-9]") # 匹配不是中文、大小写、数字的其他字符
                            for ci in range(0,len(companyInfoTitles),1):
                                if(companyInfoTitles[ci].text=='產業類別'):
                                    industry=cop.sub('', companyInfoDatas[ci].text)
                                if('統編' in companyInfoTitles[ci].text):
                                    taxNo=cop.sub('', companyInfoDatas[ci].text)
                                if(companyInfoTitles[ci].text=='負責人'):
                                    person=cop.sub('', companyInfoDatas[ci].text)
                                if(companyInfoTitles[ci].text=='資本額'):
                                    capital=cop.sub('', companyInfoDatas[ci].text)
                                if(companyInfoTitles[ci].text=='公司人數'):
                                    employeeCount=cop.sub('', companyInfoDatas[ci].text)
                            pageControl=chrome.find_element_by_id('buttonContainer')
                            address=chrome.find_element_by_class_name('job-detail-panel-content').text
                            if(address!=''):
                                address=address.split()[0]
                            infoList=chrome.find_element_by_class_name('apply-box')
                            totalJobPage=1
                            if(pageControl.text!=''):
                                startIndex=pageControl.text.find('共')+1
                                endIndex=len(pageControl.text)-1
                                totalJobPage=int(pageControl.text[startIndex:endIndex])
                            #走訪工作機會
                            for p in range(0,totalJobPage,1):
                                jobSoup= BeautifulSoup(
                                    chrome.page_source, 'html.parser')
                                jobList=jobSoup.find_all(
                                'div', {'class': 'JobDetail'})
                                for jb in jobList:
                                    jobInfos=jb.text.split()
                                    job=''
                                    for ji in range(1,len(jobInfos),1):
                                        if(jobInfos[ji]=='時薪' or jobInfos[ji]=='月薪'):
                                            break
                                        job+=jobInfos[ji]
                                    job+=" "
                                    job+=jobInfos[0]
                                    #新增job
                                    cursor.execute("insert into jobs values('"+companyName+"','"+taxNo+"','"+industry+"','"+person+"','"+capital+"','"+employeeCount+"','"+address+"','"+job+"')")
                                    mydb.commit()
                                if(totalJobPage>1 and p>0):
                                    nextPage=chrome.find_element_by_id("btnPageJump")
                                    pageJump=chrome.find_element_by_id("pageJump")
                                    pageJump.send_keys(p+1)
                                    time.sleep(2)
                                    nextPage.click()
                            #更新companyPageInfoIndex
                            companyPageInfoIndex+=1
                            cursor.execute("UPDATE CrawlerInfos set value =" +
                                        str(companyPageInfoIndex)+" where key='companyPageInfoIndex'")
                            mydb.commit()
                        
                        #重置companyPageInfoIndex
                        companyPageInfoIndex=0
                        cursor.execute("UPDATE CrawlerInfos set value =" +
                                    str(companyPageInfoIndex)+" where key='companyPageInfoIndex'")
                        mydb.commit()
                        #更新companyPageIndex
                        companyPageIndex+=1
                        cursor.execute("UPDATE CrawlerInfos set value =" +
                                    str(companyPageIndex)+" where key='companyPageIndex'")
                        mydb.commit()
                    
                    #重置companyPageIndex
                    companyPageIndex=0
                    cursor.execute("UPDATE CrawlerInfos set value =" +
                                str(companyPageIndex)+" where key='companyPageIndex'")
                    mydb.commit()
                    #更新areaIndex
                    areaIndex+=1
                    cursor.execute("UPDATE CrawlerInfos set value =" +
                                str(areaIndex)+" where key='areaIndex'")
                    mydb.commit()
                #重置companyPageIndex
                companyPageIndex=1
                cursor.execute("UPDATE CrawlerInfos set value =" +
                            str(companyPageIndex)+" where key='companyPageIndex'")
                mydb.commit()
            #重置areaIndex
            areaIndex+=0
            cursor.execute("UPDATE CrawlerInfos set value =" +
                        str(areaIndex)+" where key='areaIndex'")
            mydb.commit()
            #更新cityIndex
            cityIndex+=1
            cursor.execute("UPDATE CrawlerInfos set value =" +
                        str(cityIndex)+" where key='cityIndex'")
            mydb.commit()



            
        except Exception as e:
            print(e)
        finally:
            chrome.quit()


if __name__ == '__main__':
    Crwler104()

import requests
import sqlite3
import threading

# 定义一个函数，用于处理单个数据库文件
def process_database(db_filename):
    mydb = sqlite3.connect(db_filename)
    cursor = mydb.cursor()
    cursor.execute("SELECT d FROM FactoryData where c=''")
    rows = cursor.fetchall()
    cursor.close()
    
    for row in rows:
        url = f'https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq "{row[0]}"&$skip=0&$top=50'
        r = requests.get(url)
        
        if r.status_code == 200:
            try:
                data = r.json()
                if data:
                    Business_Accounting_NO = data[0].get('Business_Accounting_NO', '')
                    Responsible_Name = data[0].get('Responsible_Name', '')
                    
                    mydb = sqlite3.connect(db_filename)
                    cursor = mydb.cursor()
                    cursor.execute("UPDATE FactoryData SET c=? WHERE d=?", (Responsible_Name, Business_Accounting_NO))
                    mydb.commit()
                    cursor.close()
                else:
                    print(f"未找到統編 {row[0]} 的相關數據")
            except ValueError as e:
                print(f"無法解析JSON回應：{e}")
        else:
            print(f"無法獲取統編 {row[0]} 的數據，HTTP錯誤碼: {r.status_code}")
    
    mydb.close()

# 创建一个线程列表
threads = []

# 启动多个线程，每个线程处理一个数据库文件
for i in range(1, 7):
    db_filename = f"./var/db{i}.db"
    thread = threading.Thread(target=process_database, args=(db_filename,))
    threads.append(thread)
    thread.start()

# 等待所有线程完成
for thread in threads:
    thread.join()

print("所有数据库处理完成")

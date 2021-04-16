import csv
import sqlite3
# 讀取資料庫
mydb = sqlite3.connect("../var/TaxNo.db")
cursor = mydb.cursor()
cursor.execute("SELECT taxNo FROM Company WHERE companyName='千佰億彩券行'")
for row in cursor:
    print(row[0])
# 開啟 CSV 檔案
# with open('BGMOPEN1.csv', newline='', encoding="utf-8") as csvFile:
#     # 1.直接讀取：讀取 CSV 檔案內容
#     rows = csv.reader(csvFile)
#     # 2.自訂分隔符號：讀取 CSV 檔案內容
#     rows = csv.reader(csvFile, delimiter=',')
#     c = 0
    # 迴圈輸出 每一列 壹欣科技有限公司篤敬文信停車場
    # for row in rows:
    #     if(c > 1):
    #         # 更新infoPageIndex
    #         sqlStr = "INSERT INTO Company(companyName,taxNo) VALUES('" + \
    #             row[3]+"','"+row[1]+"');"
    #         cursor.execute(sqlStr)
    #         mydb.commit()
    #     c += 1

import sqlite3

# 初始化資料庫計數器
db_counter = 1

# 循環遍歷每個資料庫並清除數據
for i in range(1, 16):  # range(1, 16) 生成1至15的數字
    db_path = f"./var/db{db_counter}.db"  # 建立資料庫路徑字符串
    mydb = sqlite3.connect(db_path)  # 連接到資料庫
    cursor = mydb.cursor()
    cursor.execute("delete from CrawlerData")  # 從CrawlerData表中刪除所有記錄
    mydb.commit()  # 提交更改
    mydb.close()  # 關閉資料庫連接
    db_counter += 1  # 增加資料庫計數器

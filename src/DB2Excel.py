import sqlite3
import pandas as pd

# Create an empty DataFrame
df = pd.DataFrame()

# Loop through db1 to db5 database files
for db_number in range(1, 16):
    db_filename = f"./var/db{db_number}.db"
    mydb = sqlite3.connect(db_filename)
    
    # Read data from the database and concatenate it with the DataFrame
    query = "SELECT a,b,c,d,e FROM crawlerdata "
    temp_df = pd.read_sql_query(query, mydb)
    df = pd.concat([df, temp_df], ignore_index=True)

# Save the DataFrame to an Excel file
df.to_excel('./src/20240924.xlsx', index=False)

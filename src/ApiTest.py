import requests
r = requests.get('https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq 20828393&$skip=0&$top=50')
print(len(r.text))
print(r.json()[0]['Responsible_Name'])
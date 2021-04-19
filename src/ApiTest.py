import requests
r = requests.get('http://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq 12345678&$skip=0&$top=50')
print(len(r.text))
print(r.json()[0]['Business_Accounting_NO'])
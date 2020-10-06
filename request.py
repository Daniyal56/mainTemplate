import requests
from pprint import pprint
from docxtpl import DocxTemplate
from flask import request


user = 'GSAGROPAK- 17865'
data = requests.get(
    'http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no=12312')

# print(data.status_code)

data = data.json()


for x in data['items']:
    pprint(x)

# print('Data is now overwriting..............')
# print(x)

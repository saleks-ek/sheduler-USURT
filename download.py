import os
import requests

os.chdir('Расписания')
'''
url = 'https://bb.usurt.ru/bbcswebdav/xid-20933635_1'
response = requests.get(url)
file_Path = 'ФУПП 1 курс нечетная.xls'
'''
url = 'https://bb.usurt.ru/bbcswebdav/xid-21045128_1'
# 'https://bb.usurt.ru/bbcswebdav/institution/%D0%A0%D0%B0%D1%81%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D0%B5/%D0%97%D0%B0%D0%BE%D1%87%D0%BD%D0%B0%D1%8F%20%D1%84%D0%BE%D1%80%D0%BC%D0%B0%20%D0%BE%D0%B1%D1%83%D1%87%D0%B5%D0%BD%D0%B8%D1%8F/2024-2025%20%D1%83%D1%87%D0%B5%D0%B1%D0%BD%D1%8B%D0%B9%20%D0%B3%D0%BE%D0%B4/1%20%D1%81%D0%B5%D0%BC%D0%B5%D1%81%D1%82%D1%80/%D0%97%D0%A4%201%20%D0%BA%D1%83%D1%80%D1%81%20%D1%81%205%20%D0%BD%D0%BE%D1%8F%D0%B1%D1%80%D1%8F.xls' 
response = requests.get(url)
file_Path = 'ответ+.txt'
pass

if response.status_code == 200:
    with open(file_Path, 'wb') as file:
        file.write(response.content)
    print('File downloaded successfully')
else:
    print('Failed to download file')
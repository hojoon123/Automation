import openpyxl
import random

wb = openpyxl.Workbook()

ws = wb.active

ws.title = 'data'

ws.append(['순번', '제품명', '가격', '수량', '합계'])

name_list = ['기계식 키보드', '게이밍 마우스', '32인치 모니터', '마우스 패드']

for i in range(random.randint(5, 10)):
    name = random.choice(name_list)
    if name == '기계식 키보드':
        price = 120000
    elif name == '게이밍 마우스':
        price = 40000
    elif name == '32인치 모니터':
        price = 350000
    elif name == '마우스 패드':
        price = 20000
    
    ws.append([i + 1, name, price, random.randint(1, 5), f'=C{i+2}*D{i+2}'])
    
wb.save('03.엑셀 자동화/11번가.xlsx')
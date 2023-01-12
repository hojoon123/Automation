import openpyxl

# 새로운 엑셀 파일 생성
wb = openpyxl.Workbook()

# 현재 활성화 된 시트 선택
ws = wb.active

# Sheet 이름 변경
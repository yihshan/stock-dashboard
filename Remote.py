import openpyxl

wb = openpyxl.load_workbook('C:\Users\yihsh\Desktop\交接工作清單.xlsx')
for sheet in wb.worksheets:
    sheet.protection.sheet = False
    sheet.protection.password = None
wb.save('unlocked_file.xlsx')

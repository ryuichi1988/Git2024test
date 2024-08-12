import openpyxl

path = r"C:\Users\timaz\OneDrive\デスクトップ\1.xlsx"
file = openpyxl.load_workbook(path,data_only=True)
sheet= file.worksheets[1]
print(sheet.title)
print(sheet["H40"].value)
print(sheet["I40"].value)




import openpyxl

origem = "Planilha Julho 7 01_10 copy.xlsx"

wb = openpyxl.load_workbook(origem)
print("Todas as abas no arquivo (incluindo ocultas):")
for aba in wb.sheetnames:
    print("-", aba)

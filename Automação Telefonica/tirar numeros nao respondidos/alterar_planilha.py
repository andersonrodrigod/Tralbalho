from openpyxl import load_workbook

# Caminho do arquivo
arquivo = "Planilha JUNHO 01.10 Filtrada.xlsx"

# Carrega a planilha
wb = load_workbook(arquivo)

# Seleciona a primeira aba (ou use wb["Nome_da_aba"] se quiser algo específico)
aba = wb.active

# Altera o valor da célula A1
aba["A1"] = "COD FILIAL"

# Salva o arquivo com o mesmo nome (ou outro, se quiser preservar o original)
wb.save(arquivo)

print("✅ Cabeçalho da coluna A alterado para 'COD FILIAL'")

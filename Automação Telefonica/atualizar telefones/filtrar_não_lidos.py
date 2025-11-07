import pandas as pd

# Nome do arquivo e da aba
nome_arquivo = 'Planilha Agosto 07.11 1.xlsx'
aba = 'BASE'

# Carregar a aba 'BASE' da planilha
df = pd.read_excel(nome_arquivo, sheet_name=aba, engine='openpyxl')

# Filtrar as linhas onde a coluna 'status' está vazia
filtro = df[df['status'] == "Lida"]

# Selecionar apenas a coluna 'COD USUARIO'
codigos_filtrados = filtro[['COD USUARIO']]

# Salvar os códigos filtrados em uma nova planilha
codigos_filtrados.to_excel('codigos_filtrados.xlsx', index=False)
import pandas as pd

# Carrega a aba BASE do arquivo Excel
arquivo_entrada = 'Planilha Julho 04.11.xlsx'
df = pd.read_excel(arquivo_entrada, sheet_name="BASE")

# Lista de colunas confidenciais que devem ser removidas
colunas_confidenciais = [
    'COD FILIAL', 'COD USUARIO', 'USUARIO', 'TELEFONE',
    'contratação', 'chave', 'operador', 'tipo de contato',
    'status lida', 'atualizar contato?'
]

# Remove apenas as colunas que existem no DataFrame
df_limpo = df.drop(columns=[col for col in colunas_confidenciais if col in df.columns])

# Salva em um novo arquivo com uma aba chamada 'BASE'
df_limpo.to_excel('Planilha Julho 04.11_base.xlsx', sheet_name='BASE', index=False)

print("Arquivo limpo gerado com sucesso como 'Planilha Julho 04.11_base.xlsx'")

import pandas as pd

# Caminho do arquivo de entrada
input_file = 'planilhas/segundo_contato_geral.xlsx'

# Carregar o arquivo Excel
xls = pd.ExcelFile(input_file, engine='openpyxl')

# Ler as abas
df_segundo_contato = pd.read_excel(xls, sheet_name='segundo_contato')
df_codigos_nao_lidos = pd.read_excel(xls, sheet_name='codigos_não_lidos')
df_contatos_lidos = pd.read_excel(xls, sheet_name='contatos_lidos')

# Criar conjunto com os códigos da aba codigos_não_lidos
codigos_nao_lidos_set = set(df_codigos_nao_lidos['Codigo'])

# Verificar quais códigos da aba codigos_não_lidos estão presentes na aba segundo_contato
mask = df_segundo_contato['Codigo'].isin(codigos_nao_lidos_set)

# Selecionar as linhas da aba segundo_contato que correspondem aos códigos não lidos
rows_to_move = df_segundo_contato[mask]

# Remover essas linhas da aba segundo_contato
df_segundo_contato_filtered = df_segundo_contato[~mask]

# Adicionar essas linhas à aba contatos_lidos
df_contatos_lidos_updated = pd.concat([df_contatos_lidos, rows_to_move], ignore_index=True)

# Salvar os DataFrames atualizados em um novo arquivo Excel
output_file = 'planilhas/segundo_contato_geral_atualizado.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_segundo_contato_filtered.to_excel(writer, sheet_name='segundo_contato', index=False)
    df_codigos_nao_lidos.to_excel(writer, sheet_name='codigos_não_lidos', index=False)
    df_contatos_lidos_updated.to_excel(writer, sheet_name='contatos_lidos', index=False)

print(f"Arquivo Excel atualizado salvo como {output_file}")
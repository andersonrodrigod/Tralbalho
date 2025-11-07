import pandas as pd

# Carrega a planilha gerada com os contatos
df_contatos = pd.read_excel("Contatos_detalhados.xlsx")

# Carrega a aba BASE da planilha original
df_base = pd.read_excel("Planilha JUNHO 01.10 Filtrada.xlsx", sheet_name="BASE")

# Extrai os nomes da coluna USUARIO da planilha de contatos
usuarios_contatos = df_contatos["USUARIO"].dropna().unique()

# Filtra a BASE para manter apenas os nomes que estão na planilha de contatos
df_base_filtrada = df_base[df_base["USUARIO"].isin(usuarios_contatos)]

# Conta quantas vezes cada nome aparece
contagem = df_base_filtrada["USUARIO"].value_counts()

# Seleciona os nomes que aparecem mais de uma vez
nomes_duplicados = contagem[contagem > 1].index

# Filtra novamente a BASE para manter apenas os duplicados
df_duplicados = df_base_filtrada[df_base_filtrada["USUARIO"].isin(nomes_duplicados)]

# Seleciona as colunas desejadas
colunas_desejadas = ["COD USUARIO", "USUARIO", "TELEFONE", "COD FILIAL", "PRESTADOR"]
df_duplicados = df_duplicados[colunas_desejadas]

# Salva os resultados em uma nova planilha
df_duplicados.to_excel("Contatos_duplicados_na_detalhados.xlsx", index=False)

# Imprime o resultado
print(f"🔁 Contatos duplicados encontrados na BASE: {len(df_duplicados)}")

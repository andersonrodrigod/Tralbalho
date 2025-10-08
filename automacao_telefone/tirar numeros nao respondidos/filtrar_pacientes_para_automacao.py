import pandas as pd

# Carrega as duas abas da planilha
arquivo = "Planilha JUNHO 01.10 Filtrada.xlsx"
df_status = pd.read_excel(arquivo, sheet_name="status")
df_base = pd.read_excel(arquivo, sheet_name="BASE")

# Filtra a aba status
df_status = df_status[df_status["Status"] != "Lida"]
df_status = df_status[df_status["Respondido"] == "Não"]
df_status = df_status[df_status["Contato"].notna()]
df_status = df_status[df_status["Contato"].str.contains("Junho", na=False)]
df_status = df_status.drop_duplicates(subset="Contato", keep='last')
df_status["Contato"] = df_status["Contato"].str.replace(r"_.*", "", regex=True)

# Limpa os nomes na aba BASE também (para comparar corretamente)
df_base["USUARIO"] = df_base["USUARIO"].str.replace(r"_.*", "", regex=True)

# Filtra BASE para manter apenas os que têm Status diferente de "Lida"
df_base_filtrada = df_base[df_base["Status"] != "Lida"]

# Faz a comparação: mantém apenas os contatos que estão na BASE filtrada
contatos_validos = df_base_filtrada[df_base_filtrada["USUARIO"].isin(df_status["Contato"])]

# Contatos que não entraram
contatos_invalidos = df_status[~df_status["Contato"].isin(df_base_filtrada["USUARIO"])]

# Seleciona as colunas desejadas
colunas_desejadas = ["COD USUARIO", "USUARIO", "TELEFONE", "COD FILIAL", "PRESTADOR"]
df_resultado = contatos_validos[colunas_desejadas]

# Salva os dados encontrados em uma nova planilha
df_resultado.to_excel("Contatos_detalhados.xlsx", index=False)

# Imprime os resultados
print(f"✅ Contatos incluídos na nova planilha: {len(df_resultado)}")
print(f"❌ Contatos que não foram encontrados com Status diferente de 'Lida': {len(contatos_invalidos)}")

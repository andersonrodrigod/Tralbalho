import pandas as pd

# Carrega os usuários duplicados
df_duplicados = pd.read_excel("Contatos_duplicados_na_detalhados.xlsx")

# Carrega a aba BASE da planilha original
df_base = pd.read_excel("Planilha JUNHO 01.10 Filtrada.xlsx", sheet_name="BASE")

# Filtra a BASE para manter apenas os usuários que estão na planilha de duplicados
usuarios_duplicados = df_duplicados["USUARIO"].dropna().unique()
df_base_filtrada = df_base[df_base["USUARIO"].isin(usuarios_duplicados)]

# Agrupa por USUARIO e analisa os valores da coluna Status
usuarios_com_status_diferente = []
usuarios_com_status_em_branco = []

for usuario, grupo in df_base_filtrada.groupby("USUARIO"):
    status_unicos = grupo["Status"].dropna().unique()
    
    if len(status_unicos) == 0:
        usuarios_com_status_em_branco.append(usuario)
    elif len(status_unicos) > 1:
        usuarios_com_status_diferente.append(usuario)

# Cria DataFrames com apenas uma linha por usuário
df_status_diferente = df_base_filtrada[df_base_filtrada["USUARIO"].isin(usuarios_com_status_diferente)]
df_status_diferente = df_status_diferente.drop_duplicates(subset=["USUARIO"])

df_status_em_branco = df_base_filtrada[df_base_filtrada["USUARIO"].isin(usuarios_com_status_em_branco)]
df_status_em_branco = df_status_em_branco.drop_duplicates(subset=["USUARIO"])

# Salva os resultados em planilhas separadas
df_status_diferente.to_excel("Contatos_com_status_diferente.xlsx", index=False)
df_status_em_branco.to_excel("Contatos_com_status_em_branco.xlsx", index=False)

# Imprime os resultados
print(f"🔁 Contatos com status divergente (ex: Lida e Não Lida): {len(usuarios_com_status_diferente)}")
print(f"🧼 Contatos com todos os status em branco: {len(usuarios_com_status_em_branco)}")

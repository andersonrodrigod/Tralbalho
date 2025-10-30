import pandas as pd

# Carregar o arquivo (ajuste para .xlsx ou .csv conforme necessário)
df = pd.read_excel("achar.xlsx")  # ou pd.read_csv("achar.csv")

# Garantir que as colunas existem
required_cols = ["coluna original", "valor do resultado", "coluna buscar 1", "coluna buscar 2", "pegar valor"]
for col in required_cols:
    if col not in df.columns:
        raise ValueError(f"Coluna '{col}' não encontrada no arquivo.")

# Criar coluna para indicar qual valor foi achado
df["valor_achado"] = ""

# Loop para verificar e preencher
for i, row in df.iterrows():
    valor_original = row["coluna original"]
    resultado = None
    valor_achado = None
    # Verificar na coluna buscar 1
    match1 = df[df["coluna buscar 1"] == valor_original]
    if not match1.empty:
        resultado = match1.iloc[0]["pegar valor"]
        valor_achado = match1.iloc[0]["coluna buscar 1"]

    # Se não achou na buscar 1, verificar na buscar 2
    if resultado is None:
        match2 = df[df["coluna buscar 2"] == valor_original]
        if not match2.empty:
            resultado = match2.iloc[0]["pegar valor"]
            valor_achado = match2.iloc[0]["coluna buscar 2"]

    # Atualizar no dataframe
    if resultado is not None:
        df.at[i, "valor do resultado"] = resultado
        df.at[i, "valor_achado"] = valor_achado

# Salvar resultado
df.to_excel("achar_resultado.xlsx", index=False)
print("Processo concluído! Arquivo salvo como 'achar_resultado.xlsx'.")
import pandas as pd

# --- 1️⃣ Carregar a planilha ---
df = pd.read_excel("data_analisys.xlsx", sheet_name="respostas qr code")

# --- 2️⃣ Limpar espaços e quebras de linha nos nomes das colunas ---
df.columns = df.columns.str.strip().str.replace("\n", " ")

# --- 3️⃣ Escolher colunas que você quer manter (id_vars) ---
# Normalmente informações de identificação, filtros que quer usar no Power BI
colunas_fixas = ['sigla estado', 'sexo', 'idade']  # ajuste conforme seu arquivo

# --- 4️⃣ Derreter todas as outras colunas (value_vars) ---
colunas_para_melt = [c for c in df.columns if c not in colunas_fixas]

df_melt = df.melt(
    id_vars=colunas_fixas,
    value_vars=colunas_para_melt,
    var_name="Pergunta",
    value_name="Resposta"
)

# --- 5️⃣ Salvar resultado ---
df_melt.to_excel("resultado_melt.xlsx", index=False)

print("Transformação concluída! O arquivo 'resultado_melt.xlsx' foi criado.")

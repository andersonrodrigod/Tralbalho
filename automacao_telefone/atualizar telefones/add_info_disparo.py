import pandas as pd

# --- 1️⃣ Carregar as planilhas ---
# Substitua pelos caminhos corretos dos seus arquivos
df_principal = pd.read_excel("total_julho_minas.xlsx")
df_base = pd.read_excel("base.xlsx", sheet_name="BASE")

# --- 2️⃣ Fazer a junção das informações ---
# Aqui fazemos um merge (junção) com base nas colunas de comparação
df_resultado = pd.merge(
    df_principal,
    df_base[["COD USUARIO", "PROCEDIMENTO", "TP ATENDIMENTO"]],  # pega apenas o necessário
    left_on="Codigo",   # da planilha principal
    right_on="COD USUARIO",  # da planilha base
    how="left"          # mantém todos da principal
)

# --- 3️⃣ Remover a coluna de controle "COD USUARIO" ---
df_resultado = df_resultado.drop(columns=["COD USUARIO"])

# --- 4️⃣ Salvar a nova planilha ---
df_resultado.to_excel("resultado_com_procedimento.xlsx", index=False)

print("✅ Nova planilha criada: resultado_com_procedimento.xlsx")

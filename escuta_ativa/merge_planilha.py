import pandas as pd

# === 1. LER PLANILHA ===
arquivo = "add.xlsx"  # ajuste para o seu arquivo
df = pd.read_excel(arquivo, dtype=str)  # garante que números não percam zeros

# === 2. LIMPAR DADOS (remover espaços e caracteres não numéricos) ===
for col in ["Dado 1", "Dado 2"]:
    df[col] = df[col].str.strip().str.replace(r'\D', '', regex=True)

# === 3. Criar um dicionário Dado1 -> Nome ===
mapa_dado1_nome = dict(zip(df["Dado 1"], df["Nome"]))

# === 4. Preencher Dado 3: olhar Dado 2 e mapear para Nome da linha correspondente de Dado 1 ===
df["Dado 3"] = df["Dado 2"].map(mapa_dado1_nome)

# === 5. Exportar nova planilha ===
df.to_excel("dados_completo.xlsx", index=False)

print("✅ Planilha criada com sucesso: dados_completo.xlsx")
print(f"Foram preenchidos {df['Dado 3'].notna().sum()} registros em 'Dado 3'.")

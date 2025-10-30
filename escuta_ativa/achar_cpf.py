import pandas as pd

# === 1. LER ARQUIVO EXCEL (ajuste o nome do seu arquivo abaixo) ===
arquivo = "escuta_achar.xlsx"

# Leitura da planilha
df = pd.read_excel(arquivo)

# === 2. LIMPAR DADOS (remover espaços e garantir que sejam strings) ===
for col in ["CPFCNPJ BASE", "CPFCNPJ RESPOSTA"]:
    df[col] = (
        df[col]
        .astype(str)
        .str.strip()
        .str.replace(r'\D', '', regex=True)  # deixa só números
        .replace({'nan': '', 'None': '', '0': ''})  # trata "vazios"
    )

# === 3. FAZER O CRUZAMENTO (comparar BASE x RESPOSTA) ===
# Vamos criar uma nova tabela onde CPF/CNPJ de resposta aparece na base
correspondencias = pd.merge(
    df[["Nome Base", "CPFCNPJ BASE"]],
    df[["Nome Resposta", "CPFCNPJ RESPOSTA"]],
    left_on="CPFCNPJ BASE",
    right_on="CPFCNPJ RESPOSTA",
    how="inner"  # só mantém as que deram match
)

# === 4. FILTRAR (caso queira eliminar linhas com CPF/CNPJ vazio) ===
correspondencias = correspondencias[
    (correspondencias["CPFCNPJ BASE"] != "") &
    (correspondencias["CPFCNPJ RESPOSTA"] != "")
]

# === 5. EXPORTAR PARA NOVO EXCEL ===
correspondencias.to_excel("dados_correspondentes.xlsx", index=False)

print("✅ Planilha criada com sucesso: dados_correspondentes.xlsx")
print(f"Total de correspondências encontradas: {len(correspondencias)}")

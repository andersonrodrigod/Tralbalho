import pandas as pd

# --- 1️⃣ Carregar a planilha ---
df = pd.read_excel("dados.xlsx")

# --- 2️⃣ Remover espaços extras antes e depois do texto ---
df["Estado"] = df["Estado"].astype(str).str.strip()

# --- 3️⃣ Dicionário com os estados e suas siglas ---
estados_siglas = {
    "Acre": "AC",
    "Alagoas": "AL",
    "Amapá": "AP",
    "Amazonas": "AM",
    "Bahia": "BA",
    "Ceará": "CE",
    "Distrito Federal": "DF",
    "Espírito Santo": "ES",
    "Goiás": "GO",
    "Maranhão": "MA",
    "Mato Grosso": "MT",
    "Mato Grosso do Sul": "MS",
    "Minas Gerais": "MG",
    "Pará": "PA",
    "Paraíba": "PB",
    "Paraná": "PR",
    "Pernambuco": "PE",
    "Piauí": "PI",
    "Rio de Janeiro": "RJ",
    "Rio Grande do Norte": "RN",
    "Rio Grande do Sul": "RS",
    "Rondônia": "RO",
    "Roraima": "RR",
    "Santa Catarina": "SC",
    "São Paulo": "SP",
    "Sergipe": "SE",
    "Tocantins": "TO"
}

# --- 4️⃣ Criar nova coluna com as siglas ---
df["Sigla Estado"] = df["Estado"].map(estados_siglas)

# --- 5️⃣ Verificar se há estados não reconhecidos ---
nao_encontrados = df[df["Sigla Estado"].isna()]["Estado"].unique()
if len(nao_encontrados) > 0:
    print("⚠️ Estados não reconhecidos encontrados:")
    for estado in nao_encontrados:
        print(" -", estado)

# --- 6️⃣ Salvar o resultado ---
df.to_excel("acrescento_dados_com_siglas.xlsx", index=False)
print("✅ Nova planilha criada com a coluna de siglas.")

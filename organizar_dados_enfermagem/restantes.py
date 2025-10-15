import pandas as pd

# --- Caminho do arquivo original ---
arquivo = "plan.xlsx"  # substitua pelo seu arquivo

# --- Ler o Excel ---
df = pd.read_excel(arquivo)

# --- Garantir que os nomes sejam tratados como texto e sem espaços extras ---
df['NOME 1A'] = df['NOME 1A'].astype(str).str.strip().str.lower()
df['NOME 1B'] = df['NOME 1B'].astype(str).str.strip().str.lower()
df['NOME 2'] = df['NOME 2'].astype(str).str.strip().str.lower()

# --- Etapa 1: Remover os nomes da NOME 2 que aparecem em NOME 1A ---
nomes_1a = set(df['NOME 1A'].dropna())
df_filtrado = df[~df['NOME 2'].isin(nomes_1a)]

# --- Etapa 2: Remover também os nomes que aparecem em NOME 1B ---
nomes_1b = set(df['NOME 1B'].dropna())
df_filtrado = df_filtrado[~df_filtrado['NOME 2'].isin(nomes_1b)]

# --- Resultado final: somente os nomes de NOME 2 que não foram encontrados ---
resultado = df_filtrado[['NOME 2']].drop_duplicates().reset_index(drop=True)

# --- Salvar o resultado em um novo Excel ---
resultado.to_excel("nomes_nao_encontrados.xlsx", index=False)

print("✅ Arquivo 'nomes_nao_encontrados.xlsx' gerado com sucesso!")

import pandas as pd
import re
import unicodedata

# ====== CONFIGURAÇÕES ======
arquivo = "plan.xlsx"             # caminho do arquivo
aba = "BASE"                      # nome da aba
coluna_original = "DIFICULDADE"    # nome da coluna com os textos
saida = "planilha_frequencia_numerada.xlsx"
# ============================

# Remove acentos e coloca tudo em minúsculo
def strip_accents_lower(s):
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

# Extrai a primeira palavra útil
def primeira_palavra(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip()
    token = re.split(r'[:;,\/\-\(\)\[\]]+|\s+', texto, maxsplit=1)[0]
    token = token.strip()
    return strip_accents_lower(token)

# Mapeia o token para número
def mapa_para_numero(token):
    if token == "":
        return pd.NA

    # ====== Regras ======
    if token.startswith("nunca") or token.startswith("jamais"):
        return 1
    if token.startswith("rar") or "as vezes" in token or "vezes" in token:
        return 2
    if token.startswith("freq") or "muito" in token or "constant" in token:
        return 3
    if token.startswith("sempre") or token.startswith("toda"):
        return 4
    # fallback
    return 4

# Labels correspondentes
label_map = {
    1: "1 - Nunca",
    2: "2 - Raramente",
    3: "3 - Frequentemente",
    4: "4 - Sempre"
}

# ====== Execução ======
df = pd.read_excel(arquivo, sheet_name=aba)

# Detecta coluna automaticamente, se necessário
orig = coluna_original
if orig not in df.columns:
    candidates = [c for c in df.columns if any(k in c.lower() for k in ("freq", "vez", "habito", "period", "resposta", "observa"))]
    if candidates:
        coluna = candidates[0]
        print(f"Atenção: coluna '{orig}' não encontrada. Usando '{coluna}' (auto-detect).")
    else:
        poss = [c for c in df.columns if c.lower() != "usuario"]
        if not poss:
            raise ValueError("Não encontrei uma coluna válida para processar.")
        coluna = poss[0]
        print(f"Atenção: coluna '{orig}' não encontrada. Usando '{coluna}' como fallback.")
else:
    coluna = orig

# Processa a coluna
serie_primeira = df[coluna].apply(primeira_palavra)
serie_num = serie_primeira.apply(mapa_para_numero)
serie_label = serie_num.map(label_map)

# Insere as novas colunas ao lado da original
idx = df.columns.get_loc(coluna)
df.insert(idx+1, f"{coluna}_PRIMEIRA", serie_primeira)
df.insert(idx+2, f"{coluna}_NUM", serie_num)
df.insert(idx+3, f"{coluna}_LABEL", serie_label)

# Salva resultado
df.to_excel(saida, index=False)

# Relatório
print("✅ Arquivo salvo em:", saida)
print("\nAmostra (original -> primeira -> num -> label):")
print(df[[coluna, f"{coluna}_PRIMEIRA", f"{coluna}_NUM", f"{coluna}_LABEL"]].head(15))

print("\nContagem por número (1..4):")
print(df[f"{coluna}_NUM"].value_counts(dropna=False).sort_index())

import pandas as pd
import re
import unicodedata

# ====== CONFIGURAÇÕES ======
arquivo = "plan.xlsx"             # caminho do arquivo
aba = "BASE"                      # nome da aba
coluna_original = "DIFICULDADE"       # nome da coluna que contém os textos
saida = "planilha_contatos_numerada.xlsx"
# ============================

# Função para remover acentos e deixar minúsculo
def strip_accents_lower(s):
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

# Extrai a primeira "palavra" útil
def primeira_palavra(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip()
    token = re.split(r'[:;,\/\-\(\)\[\]]+|\s+', texto, maxsplit=1)[0]
    token = token.strip()
    return strip_accents_lower(token)

# Mapeamento do token para número
def mapa_para_numero(token):
    if token == "":
        return pd.NA

    # ====== Regras principais ======
    if token.startswith("email") or token.startswith("e-mail") or "mail" in token:
        return 1
    if token.startswith("tel") or token.startswith("fone") or "numero" in token:
        return 2
    if token.startswith("whats") or token.startswith("zap"):
        return 3
    if token.startswith("rede") or token.startswith("insta") or token.startswith("face") or "tiktok" in token or "x " in token:
        return 4
    if token.startswith("outro"):
        return 5
    # fallback: considerar como "Outros"
    return 5

# Labels correspondentes
label_map = {
    1: "1 - E-mail",
    2: "2 - Telefone",
    3: "3 - WhatsApp",
    4: "4 - Redes Sociais",
    5: "5 - Outros"
}

# ====== Execução ======
df = pd.read_excel(arquivo, sheet_name=aba)

# Verifica se a coluna existe ou tenta detectar automaticamente
orig = coluna_original
if orig not in df.columns:
    candidates = [c for c in df.columns if any(k in c.lower() for k in ("contato", "comunic", "canal", "resposta", "observa"))]
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

# Extrai as colunas processadas
serie_primeira = df[coluna].apply(primeira_palavra)
serie_num = serie_primeira.apply(mapa_para_numero)
serie_label = serie_num.map(label_map)

# Insere novas colunas ao lado da original
idx = df.columns.get_loc(coluna)
df.insert(idx+1, f"{coluna}_PRIMEIRA", serie_primeira)
df.insert(idx+2, f"{coluna}_NUM", serie_num)
df.insert(idx+3, f"{coluna}_LABEL", serie_label)

# Salva resultado
df.to_excel(saida, index=False)

# Relatório rápido
print("✅ Arquivo salvo em:", saida)
print("\nAmostra (original -> primeira -> num -> label):")
print(df[[coluna, f"{coluna}_PRIMEIRA", f"{coluna}_NUM", f"{coluna}_LABEL"]].head(15))

print("\nContagem por número (1..5):")
print(df[f"{coluna}_NUM"].value_counts(dropna=False).sort_index())

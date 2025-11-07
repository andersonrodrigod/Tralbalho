import pandas as pd
import re
import unicodedata

# ====== CONFIGURE AQUI ======
arquivo = "plan" \
".xlsx"         # caminho do arquivo
aba = "BASE"                      # nome da aba
coluna_original = "DIFICULDADE"   # nome da coluna que contém os textos
saida = "planilha_organizada_numerada.xlsx"
# =============================

# Função para remover acentos e lowercase
def strip_accents_lower(s):
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

# Função que retorna a primeira "palavra" útil da célula
def primeira_palavra(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip()
    # split por : , ; / - ( ) ou espaço - pegamos o primeiro token
    token = re.split(r'[:;,\/\-\(\)\[\]]+|\s+', texto, maxsplit=1)[0]
    token = token.strip()
    return strip_accents_lower(token)

# Mapeamento do token para número
def mapa_para_numero(token):
    if token == "":
        return pd.NA  # vazio -> mantém NA (pode alterar para 4 se preferir)
    # regras por prefixo (sem acento e em minúsculas)
    if token.startswith("nao") or token.startswith("sem") or token.startswith("nenhum") or token.startswith("nenhuma"):
        return 1
    if token.startswith("comunic"):
        return 2
    if token.startswith("pagament") or token.startswith("boleto") or token.startswith("cartao"):
        return 3
    if token.startswith("outro"):
        return 4
    # fallback: considerar como "Outros"
    return 4

# labels
label_map = {
    1: "1 - Não tenho dificuldades",
    2: "2 - Comunicação",
    3: "3 - Pagamento",
    4: "4 - Outros"
}

# ====== Execução ======
df = pd.read_excel(arquivo, sheet_name=aba)

# Auto-detect: se coluna_original não existir, tenta escolher uma coluna candidata
orig = coluna_original
if orig not in df.columns:
    candidates = [c for c in df.columns if any(k in c.lower() for k in ("dific", "problema", "motivo", "coment", "resposta", "observa"))]
    if candidates:
        coluna = candidates[0]
        print(f"Atenção: coluna '{orig}' não encontrada. Usando '{coluna}' (auto-detect).")
    else:
        # fallback para primeira coluna que não seja USUARIO
        poss = [c for c in df.columns if c.lower() != "usuario"]
        if not poss:
            raise ValueError("Não encontrei uma coluna válida para processar.")
        coluna = poss[0]
        print(f"Atenção: coluna '{orig}' não encontrada. Usando '{coluna}' como fallback.")
else:
    coluna = orig

# Extrair primeira palavra, número e label
serie_primeira = df[coluna].apply(primeira_palavra)
serie_num = serie_primeira.apply(mapa_para_numero)
serie_label = serie_num.map(label_map)

# Inserir as colunas ao lado da original (após a coluna original)
idx = df.columns.get_loc(coluna)
df.insert(idx+1, f"{coluna}_PRIMEIRA", serie_primeira)
df.insert(idx+2, f"{coluna}_NUM", serie_num)
df.insert(idx+3, f"{coluna}_LABEL", serie_label)

# Salvar
df.to_excel(saida, index=False)

# Relatórios rápidos
print("✅ Salvo em:", saida)
print("\nAmostra (original -> primeira -> num -> label):")
print(df[[coluna, f"{coluna}_PRIMEIRA", f"{coluna}_NUM", f"{coluna}_LABEL"]].head(15))

print("\nContagem por número (1..4):")
print(df[f"{coluna}_NUM"].value_counts(dropna=False).sort_index())

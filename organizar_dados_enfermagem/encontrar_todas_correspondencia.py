import pandas as pd

# --- Caminho do arquivo ---
arquivo = "novo.xlsx"  # substitua pelo seu arquivo

# --- Ler a planilha ---
df = pd.read_excel(arquivo)

# --- Palavras a ignorar ---
ignorar = {"de", "da", "do", "dos"}

# --- Função para normalizar e remover palavras irrelevantes ---
def limpar_nome(nome):
    partes = nome.strip().lower().split()
    return [p for p in partes if p not in ignorar]

# --- Listas de resultados ---
resultados = []

print("🔍 Iniciando comparação...")

# --- Loop principal ---
for i, nome in enumerate(df["Nome 1"].astype(str), 1):
    nome_limpo = limpar_nome(nome)
    matchs_3 = []
    matchs_2 = []
    matchs_1 = []

    # --- Verificar 3 palavras ---
    if len(nome_limpo) >= 3:
        chave3 = " ".join(nome_limpo[:3])
        for base in df["Nome 2"].astype(str):
            base_limpa = " ".join(limpar_nome(base))
            if base_limpa.startswith(chave3):
                matchs_3.append(base)

    # --- Verificar 2 palavras ---
    if len(nome_limpo) >= 2:
        chave2 = " ".join(nome_limpo[:2])
        for base in df["Nome 2"].astype(str):
            base_limpa = " ".join(limpar_nome(base))
            if base_limpa.startswith(chave2):
                matchs_2.append(base)

    # --- Verificar 1 palavra ---
    if len(nome_limpo) >= 1:
        chave1 = nome_limpo[0]
        for base in df["Nome 2"].astype(str):
            base_limpa = limpar_nome(base)
            if chave1 in base_limpa:
                matchs_1.append(base)

    resultados.append({
        "Nome 1": nome,
        "Match 3 Palavras": ", ".join(matchs_3) if matchs_3 else "",
        "Match 2 Palavras": ", ".join(matchs_2) if matchs_2 else "",
        "Match 1 Palavra": ", ".join(matchs_1) if matchs_1 else ""
    })

    if i % 10 == 0 or i == len(df):
        print(f"🔸 Processados {i}/{len(df)} nomes...")

# --- Criar DataFrame final ---
df_final = pd.DataFrame(resultados)

# --- Salvar em Excel ---
saida = "final_todos_matchs.xlsx"
df_final.to_excel(saida, index=False)

print("\n✅ Processamento concluído!")
print(f"📄 Arquivo '{saida}' criado com {len(df_final)} linhas.")
print("➡ Cada linha contém as correspondências de 3, 2 e 1 palavra nas colunas ao lado.")

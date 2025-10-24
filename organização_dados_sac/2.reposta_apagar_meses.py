import pandas as pd
import warnings

warnings.simplefilter("ignore", UserWarning)

# Caminho do arquivo de entrada
caminho_arquivo = "planilhas/detalhamento_de_pesquisa_eletivo.xlsx"
meses_para_remover = ["Junho de 2025", "Julho de 2025"]

# Ler todas as abas do Excel
planilhas = pd.read_excel(caminho_arquivo, sheet_name=None)

# Novo dicionário para armazenar as abas limpas
planilhas_limpa = {}

for nome_aba, df in planilhas.items():
    print(f"🔄 Processando aba: {nome_aba}")

    # Converte tudo pra string temporariamente (evita erro com NaN)
    df_str = df.astype(str)

    # Cria uma máscara: True se a linha NÃO contém nenhum dos meses
    mascara = ~df_str.apply(
        lambda col: col.str.contains("|".join(meses_para_remover), na=False)
    ).any(axis=1)

    # Aplica o filtro
    df_limpo = df.loc[mascara].copy()

    planilhas_limpa[nome_aba] = df_limpo
    print(f"✅ {len(df) - len(df_limpo)} linhas removidas da aba '{nome_aba}'")

# Salvar todas as abas novamente
saida = "detalhamento_de_pesquisa_eletivo_limpo.xlsx"
with pd.ExcelWriter(saida, engine="openpyxl") as writer:
    for nome_aba, df in planilhas_limpa.items():
        df.to_excel(writer, sheet_name=nome_aba, index=False)

print("💾 Arquivo salvo com sucesso!")

import os
import pandas as pd

# --- 1️⃣ Caminho principal ---
pasta_principal = r"C:\Users\anderson.dossantos\Desktop\dev\Tralbaho\gestar"

# --- 2️⃣ Lista de meses ---
meses = ["dez", "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out"]

# --- 3️⃣ Função para detectar onde começa a tabela ---
def encontrar_inicio_tabela(caminho):
    df_raw = pd.read_excel(caminho, header=None)
    for i, linha in df_raw.iterrows():
        if "Filial" in linha.values or "Unidades" in linha.values:
            return i
    return None

# --- 4️⃣ Dicionário para armazenar colunas de cada arquivo ---
colunas_por_arquivo = {}

# --- 5️⃣ Ler arquivos e guardar colunas ---
for mes in meses:
    pasta_mes = os.path.join(pasta_principal, mes)
    if not os.path.exists(pasta_mes):
        print(f"⚠️ Pasta {mes} não encontrada.")
        continue

    for arquivo in os.listdir(pasta_mes):
        if arquivo.lower().endswith(".xlsx") and "data 1" in arquivo.lower():
            caminho_arquivo = os.path.join(pasta_mes, arquivo)

            linha_inicio = encontrar_inicio_tabela(caminho_arquivo)
            if linha_inicio is None:
                print(f"❌ {mes.upper()}: cabeçalho não encontrado ({arquivo})")
                continue

            df = pd.read_excel(caminho_arquivo, skiprows=linha_inicio)
            colunas_por_arquivo[f"{mes}/{arquivo}"] = list(df.columns)

# --- 6️⃣ Comparar com o primeiro arquivo como referência ---
arquivos = list(colunas_por_arquivo.keys())

if not arquivos:
    print("\n⚠️ Nenhum arquivo encontrado para verificação.")
else:
    ref_arquivo = arquivos[0]
    ref_colunas = set(colunas_por_arquivo[ref_arquivo])

    print(f"\n📘 Arquivo de referência: {ref_arquivo}")
    print(f"🧩 Total de colunas: {len(ref_colunas)}\n")

    # --- 7️⃣ Verificar diferenças ---
    for nome, colunas in colunas_por_arquivo.items():
        colunas_set = set(colunas)
        faltando = ref_colunas - colunas_set
        extras = colunas_set - ref_colunas

        if not faltando and not extras:
            print(f"✅ {nome} → Colunas OK ({len(colunas)} colunas)")
        else:
            print(f"\n⚠️ {nome} → Diferenças encontradas:")
            if faltando:
                print(f"   ❌ Faltando: {', '.join(faltando)}")
            if extras:
                print(f"   ⚠️ Extras: {', '.join(extras)}")

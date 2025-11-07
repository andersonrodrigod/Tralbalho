import pandas as pd
import warnings

# Ignorar avisos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

arquivos = [
    "julho_agosto/detalhamento_de_pesquisa_escuta_ativa_vendedores_21.07 - 20.08.xlsx",
    "agosto_setembro/detalhamento_de_pesquisa_escuta_ativa_vendedores 21.08 - 20.09.xlsx",
    "setembro_outubro/detalhamento_de_pesquisa_escuta_ativa_vendedores 21.09 - 20.10.xlsx"
]

abas_unidas = {}

abas_primeiro = pd.read_excel(arquivos[0], sheet_name=None)

print("\n===== CONTAGEM DE LINHAS POR ABA E ARQUIVO =====\n")

for nome_aba in abas_primeiro.keys():
    lista_abas = []
    total_linhas_esperado = 0

    print(f"\n📄 ABA: {nome_aba}")

    for arquivo in arquivos:
        df = pd.read_excel(arquivo, sheet_name=nome_aba)
        linhas = len(df)
        total_linhas_esperado += linhas
        print(f"  - {arquivo.split('/')[-1]} → {linhas} linhas")
        lista_abas.append(df)

    df_concatenado = pd.concat(lista_abas, ignore_index=True)
    abas_unidas[nome_aba] = df_concatenado

    print(f"  ➕ Total esperado: {total_linhas_esperado}")
    print(f"  ✅ Total após concatenação: {len(df_concatenado)}")

with pd.ExcelWriter("unificado_todas_as_abas.xlsx", engine="openpyxl") as writer:
    for nome_aba, df in abas_unidas.items():
        df.to_excel(writer, sheet_name=nome_aba, index=False)

print("\n✅ Todas as abas foram unidas e salvas com sucesso!")

import pandas as pd
arquivo = "junho.xlsx"

abas_desejadas = ["BASE", "status"]

todas_as_abas = pd.read_excel(arquivo, sheet_name=None)

abas_filtradas = {nome: todas_as_abas[nome] for nome in abas_desejadas if nome in todas_as_abas}


with pd.ExcelWriter("junho_abas_desejadas.xlsx") as writter:
    for nome, aba in abas_filtradas.items():
        aba.to_excel(writter, sheet_name=nome, index=False)



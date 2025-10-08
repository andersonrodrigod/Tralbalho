import pandas as pd

arquivo = "planilhas/junho_abas_desejadas.xlsx"

abas = pd.read_excel(arquivo, sheet_name=None)

base = abas["BASE"]

status = abas["status"]

base.columns = base.columns.str.upper()
status.columns = status.columns.str.upper()

comparacao = base.merge(status, on="TELEFONE", suffixes=('_aba1', "_aba2"))


comparacao["CHAVE_AUX"] = comparacao["CHAVE"].str.split("_").str[0]
comparacao["CONTATO_AUX"] = comparacao["CONTATO"].str.split("_").str[0]

diferentes = comparacao[comparacao["CHAVE_AUX"] != comparacao["CONTATO_AUX"]]

linhas_txt = []
for _, linha in diferentes.iterrows():
    texto = f"Telefone {linha['TELEFONE']} - Aba1: {linha['CHAVE']} | Aba2: {linha['CONTATO']}"
    linhas_txt.append(texto)
    


with open("usuarios_diferentes.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(linhas_txt))












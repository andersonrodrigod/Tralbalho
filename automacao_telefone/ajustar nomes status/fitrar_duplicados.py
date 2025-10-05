import pandas as pd

arqiuvo = "planilhas/junho_abas_desejadas.xlsx"

abas = pd.read_excel(arqiuvo, sheet_name=None)

base = abas["BASE"]


duplicados = base[base.duplicated(subset=["TELEFONE"], keep=False)].copy()

usuarios_duplicados = duplicados[["USUARIO", "TELEFONE"]]

usuarios_duplicados.to_excel("usuarios_duplicados.xlsx", index=False)

"""with open("usuarios_duplicados.txt", "w", encoding="utf-8") as f:
    for _, linha in usuarios_duplicados.iterrows():
        f.write(f"{linha["USUARIO"]} - {linha["TELEFONE"]} \n")"""

print(f"Foram encontrados {len(usuarios_duplicados)} telefones duplicados.")

























import pandas as pd


def contar_status(df):
    return {
        "Lidas": df[df["Status"] == "Lida"].shape[0],
        "Não quis": df[df["Status"] == "Não quis"].shape[0],
        "Óbito": df[df["Status"] == "Óbito"].shape[0],
        "Sem resultado": df[df["Status"].isna()].shape[0]
    }

def contar_status_resposta(df):
    return {
        "p1": df[(df["Status"] == "Lida") & (df["p1"].notna())].shape[0],
        "p2": df[(df["Status"] == "Lida") & (df["p2"].notna())].shape[0],
        "p3": df[(df["Status"] == "Lida") & (df["p3"].notna())].shape[0],
        "p4": df[(df["Status"] == "Lida") & (df["p4"].notna())].shape[0],
        "p5": df[(df["Status"] == "Lida") & (df["p5"].notna())].shape[0],
        "p6": df[(df["Status"] == "Lida") & (df["p6"].notna())].shape[0],
    }

def contar_elogio_queixas_geral(df):
    tabelas = []

    # Percorre os blocos de 1 a 6
    for i in range(1, 7):
        grupo_col = f"GRUPO-{i}"
        motivo_col = f"MOTIVO-{i}"
        tipo_col = f"ELOGIO OU QUEIXA-{i}"

        # Verifica se as três colunas existem na planilha
        if all(col in df.columns for col in [grupo_col, motivo_col, tipo_col]):
            # Agrupa por grupo, motivo e tipo (ELOGIO/QUEIXA)
            df_filtrado = df.dropna(subset=[grupo_col, motivo_col, tipo_col])
            if not df_filtrado.empty:
                tabela = (
                    df_filtrado.groupby([grupo_col, motivo_col, tipo_col])
                    .size()
                    .reset_index(name="Quantidade")
                )
                tabela.columns = ["Grupo", "Motivo", "Tipo", "Quantidade"]
                tabelas.append(tabela)

    # Junta todas as tabelas em uma só
    if not tabelas:
        return pd.DataFrame()  # Caso não haja dados

    tabela_final = pd.concat(tabelas, ignore_index=True)

    # Faz o pivot para ter ELOGIO/QUEIXA como colunas
    tabela_resumo = tabela_final.pivot_table(
        index=["Grupo", "Motivo"],
        columns="Tipo",
        values="Quantidade",
        aggfunc="sum",
        fill_value=0
    )

    return tabela_resumo


def contar_elogio(df):
    tabelas = []

    # Percorre os 6 blocos de colunas
    for i in range(1, 7):
        grupo_col = f"GRUPO-{i}"
        motivo_col = f"MOTIVO-{i}"
        tipo_col = f"ELOGIO OU QUEIXA-{i}"

        # Confere se as colunas existem
        if all(col in df.columns for col in [grupo_col, motivo_col, tipo_col]):
            # Filtra apenas ELOGIO/QUEIXA
            df_filtrado = df[df[tipo_col] == "ELOGIO"]

            if not df_filtrado.empty:
                # Agrupa e cria tabela
                tabela = df_filtrado.groupby([grupo_col, motivo_col]).size().reset_index(name="Quantidade")
                tabela.columns = ["Grupo", "Motivo", "Quantidade"]
                tabelas.append(tabela)

    # Junta todas as tabelas em uma só
    if not tabelas:
        return pd.DataFrame()  # caso não haja nenhuma

    tabela_final = pd.concat(tabelas, ignore_index=True)

    # Faz o pivot (transforma em tabela larga)
    tabela_resumo = tabela_final.pivot_table(
        index="Grupo",
        columns="Motivo",
        values="Quantidade",
        aggfunc="sum",
        fill_value=0
    )

    # 🔹 Garante que todos os grupos e motivos apareçam, mesmo com 0
    todos_grupos = []
    todos_motivos = []
    for i in range(1, 7):
        if f"GRUPO-{i}" in df.columns:
            todos_grupos.extend(df[f"GRUPO-{i}"].dropna().unique())
        if f"MOTIVO-{i}" in df.columns:
            todos_motivos.extend(df[f"MOTIVO-{i}"].dropna().unique())

    # Remove duplicados e mantém ordem
    todos_grupos = sorted(set(todos_grupos))
    todos_motivos = sorted(set(todos_motivos))

    # Reindexa com 0 para tudo que não apareceu
    tabela_resumo = tabela_resumo.reindex(index=todos_grupos, columns=todos_motivos, fill_value=0)

    return tabela_resumo

def contar_queixas(df):
    tabelas = []

    # Percorre os 6 blocos de colunas
    for i in range(1, 7):
        grupo_col = f"GRUPO-{i}"
        motivo_col = f"MOTIVO-{i}"
        tipo_col = f"ELOGIO OU QUEIXA-{i}"

        # Confere se as colunas existem
        if all(col in df.columns for col in [grupo_col, motivo_col, tipo_col]):
            # Filtra apenas ELOGIO/QUEIXA
            df_filtrado = df[df[tipo_col] == "QUEIXA"]

            if not df_filtrado.empty:
                # Agrupa e cria tabela
                tabela = df_filtrado.groupby([grupo_col, motivo_col]).size().reset_index(name="Quantidade")
                tabela.columns = ["Grupo", "Motivo", "Quantidade"]
                tabelas.append(tabela)

    # Junta todas as tabelas em uma só
    if not tabelas:
        return pd.DataFrame()  # caso não haja nenhuma

    tabela_final = pd.concat(tabelas, ignore_index=True)

    # Faz o pivot (transforma em tabela larga)
    tabela_resumo = tabela_final.pivot_table(
        index="Grupo",
        columns="Motivo",
        values="Quantidade",
        aggfunc="sum",
        fill_value=0
    )

    # 🔹 Garante que todos os grupos e motivos apareçam, mesmo com 0
    todos_grupos = []
    todos_motivos = []
    for i in range(1, 7):
        if f"GRUPO-{i}" in df.columns:
            todos_grupos.extend(df[f"GRUPO-{i}"].dropna().unique())
        if f"MOTIVO-{i}" in df.columns:
            todos_motivos.extend(df[f"MOTIVO-{i}"].dropna().unique())

    # Remove duplicados e mantém ordem
    todos_grupos = sorted(set(todos_grupos))
    todos_motivos = sorted(set(todos_motivos))

    # Reindexa com 0 para tudo que não apareceu
    tabela_resumo = tabela_resumo.reindex(index=todos_grupos, columns=todos_motivos, fill_value=0)

    return tabela_resumo


def contar_elogios_queixas(df):
    tabelas = []

    # Percorre os 6 blocos de colunas
    for i in range(1, 7):
        grupo_col = f"GRUPO-{i}"
        motivo_col = f"MOTIVO-{i}"
        tipo_col = f"ELOGIO OU QUEIXA-{i}"

        # Confere se as colunas existem
        if all(col in df.columns for col in [grupo_col, motivo_col, tipo_col]):
            # Filtra apenas ELOGIO/QUEIXA
            df_filtrado = df[df[tipo_col] == "ELOGIO/QUEIXA"]

            if not df_filtrado.empty:
                # Agrupa e cria tabela
                tabela = df_filtrado.groupby([grupo_col, motivo_col]).size().reset_index(name="Quantidade")
                tabela.columns = ["Grupo", "Motivo", "Quantidade"]
                tabelas.append(tabela)

    # Junta todas as tabelas em uma só
    if not tabelas:
        return pd.DataFrame()  # caso não haja nenhuma

    tabela_final = pd.concat(tabelas, ignore_index=True)

    # Faz o pivot (transforma em tabela larga)
    tabela_resumo = tabela_final.pivot_table(
        index="Grupo",
        columns="Motivo",
        values="Quantidade",
        aggfunc="sum",
        fill_value=0
    )

    # 🔹 Garante que todos os grupos e motivos apareçam, mesmo com 0
    todos_grupos = []
    todos_motivos = []
    for i in range(1, 7):
        if f"GRUPO-{i}" in df.columns:
            todos_grupos.extend(df[f"GRUPO-{i}"].dropna().unique())
        if f"MOTIVO-{i}" in df.columns:
            todos_motivos.extend(df[f"MOTIVO-{i}"].dropna().unique())

    # Remove duplicados e mantém ordem
    todos_grupos = sorted(set(todos_grupos))
    todos_motivos = sorted(set(todos_motivos))

    # Reindexa com 0 para tudo que não apareceu
    tabela_resumo = tabela_resumo.reindex(index=todos_grupos, columns=todos_motivos, fill_value=0)

    return tabela_resumo





def contar_respostas(df):
    perguntas = ["p1", "p2", "p3", "p4", "p5", "p6"]
    resultado = {}

    for pergunta in perguntas:
        resultado[pergunta] = {
            str(nota): df[df[pergunta] == nota].shape[0]
            for nota in range(1, 6)
        }

    return resultado


















# CÓDIGOS MORTOS

"""def contar_queixas(df):
    tabelas = []

    for i in range(1, 7):  # GRUPO-1 até GRUPO-6
        grupo_col = f"GRUPO-{i}"
        motivo_col = f"MOTIVO-{i}"
        tipo_col = f"ELOGIO OU QUEIXA-{i}"

        if grupo_col in df.columns and motivo_col in df.columns and tipo_col in df.columns:
            df_elogio = df[df[tipo_col] == "QUEIXA"]
            tabela = df_elogio.groupby([grupo_col, motivo_col]).size().reset_index(name="Quantidade")
            tabela.columns = ["Grupo", "Motivo", "Quantidade"]
            tabelas.append(tabela)

    # Junta todas as tabelas em uma só
    tabela_final = pd.concat(tabelas, ignore_index=True)

    # Agrupa novamente para somar os elogios por grupo e motivo
    tabela_resumo = tabela_final.groupby(["Grupo", "Motivo"]).sum().unstack(fill_value=0)

    return tabela_resumo"""



"""def contar_elogio(df):
    df_elogio = df[df["ELOGIO OU QUEIXA-1"] == "ELOGIO"]
    tabela = df_elogio.groupby(["GRUPO-1", "MOTIVO-1"]).size().unstack(fill_value=0)
    return tabela"""


"""
def contar_queixas(df):
    df_elogio = df[df["ELOGIO OU QUEIXA-1"] == "QUEIXA"]
    tabela = df_elogio.groupby(["GRUPO-1", "MOTIVO-1"]).size().unstack(fill_value=0)
    return tabela"""
    
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
    tabela = df.groupby(["GRUPO-1", "MOTIVO-1", "ELOGIO OU QUEIXA-1"]).size().unstack(fill_value=0)
    return tabela

def contar_elogio(df):
    df_elogio = df[df["ELOGIO OU QUEIXA-1"] == "ELOGIO"]
    tabela = df_elogio.groupby(["GRUPO-1", "MOTIVO-1"]).size().unstack(fill_value=0)
    return tabela


def contar_queixas(df):
    df_elogio = df[df["ELOGIO OU QUEIXA-1"] == "QUEIXA"]
    tabela = df_elogio.groupby(["GRUPO-1", "MOTIVO-1"]).size().unstack(fill_value=0)
    return tabela
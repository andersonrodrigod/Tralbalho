import pandas as pd



def load_data(path="Planilha Julho 04.11.xlsx"):
    df = pd.read_excel(path, sheet_name="BASE")
    df = df.drop_duplicates()

    return df










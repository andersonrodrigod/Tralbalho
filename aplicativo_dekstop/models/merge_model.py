import pandas as pd
import os

class MergeModel:
    def __init__(self, file1, file2):
        self.file1 = file1
        self.file2 = file2

    def juntar(self):
        df1 = pd.read_excel(self.file1)
        df2 = pd.read_excel(self.file2)
        df_final = pd.concat([df1, df2])
        
        pasta = os.path.dirname(self.file1)
        caminho_saida = os.path.join(pasta, "arquivos_unidos.xlsx")
        df_final.to_excel(caminho_saida, index=False)
        return caminho_saida

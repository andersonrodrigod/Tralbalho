import pandas as pd
from utils.abas_padrao import abas_eletivo, abas_internacao

class DetalhamentoController:
    def __init__(self):
        pass
    

    def ajustar_abas(self, caminho_arquivo: str, tipo: str):
        """
        Lê um arquivo Excel e ajusta os nomes das abas para o padrão definido.
        Retorna um dicionário com {nome_aba_padrao: DataFrame}.
        
        Parâmetros:
        - caminho_arquivo: str → caminho do arquivo Excel
        - tipo: str → "eletivo" ou "internacao"
        """

        # 1️⃣ Escolhe o padrão de abas correto
        abas_padrao = abas_eletivo if tipo == "eletivo" else abas_internacao

        # 2️⃣ Carrega o arquivo Excel inteiro (não lemos ainda as abas)
        xls = pd.ExcelFile(caminho_arquivo)

        # 3️⃣ Cria o dicionário onde vamos armazenar os DataFrames
        # A chave será o nome da aba (ajustado ou original)
        # O valor será o DataFrame da aba
        df_dict = {}
        #abas_nao_ajustadas = [] 

        # 4️⃣ Itera por todas as abas do arquivo
        for aba in xls.sheet_names:
            #print(f"\n➡️ Verificando aba: {aba}")
            aba_encontrada = None  # inicialmente nenhuma aba padronizada encontrada

            # 5️⃣ Compara cada aba do arquivo com os nomes padrões
            for chave, nome_padrao in abas_padrao.items():
                #print(f"   🔹 Comparando com chave='{chave}' e nome_padrao='{nome_padrao}'")
                if nome_padrao == aba:  # << se bater
                    aba_encontrada = chave
                    print(f"   ✅ Aba encontrada! Será renomeada para: {aba_encontrada}")
                    break 

            # 6️⃣ Se encontrou algum padrão, lê a aba e adiciona ao dicionário com o nome padrão
            if aba_encontrada:
                df = pd.read_excel(caminho_arquivo, sheet_name=aba)
                df_dict[aba_encontrada] = df  # chave = nome padrão, valor = DataFrame
            else:
                # 7️⃣ Caso não bata com nenhum padrão, mantém o nome original
                #print(f"   ⚠️ Nenhum padrão corresponde. Mantendo nome original: {aba}")
                df_dict[aba] = pd.read_excel(caminho_arquivo, sheet_name=aba)


        #if abas_nao_ajustadas:
            #print("\n🚨 Abas não ajustadas (não bateram com nenhum padrão):")
            #for aba in abas_nao_ajustadas:
                #print(f" - {aba}")

        # 8️⃣ Retorna o dicionário com todas as abas ajustadas
        return df_dict

    def juntar_abas(self, lista_dfs):
        """
        Recebe uma lista de dicionários de DataFrames (um por arquivo),
        e concatena as abas iguais.
        Retorna um único dicionário {nome_aba_padrao: DataFrame concatenado}.
        """
        
        resultado = {}

        for dfs in lista_dfs:
            for aba, df in dfs.items():
                if aba in resultado:
                   resultado[aba] = pd.concat([resultado[aba], df], ignore_index=True)
                else:
                    resultado[aba] = df.copy()
        return resultado
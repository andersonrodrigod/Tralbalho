import pandas as pd

class ConcatenarController:
    def __init__(self):
        pass

    def processar_arquivos(self, lista_arquivos):
        """
        Recebe uma lista de caminhos de arquivos Excel e concatena 
        as abas de mesmo nome entre todos os arquivos.
        
        Retorna:
            dict[str, pd.DataFrame]: {nome_aba: DataFrame concatenado}
        """
        if not lista_arquivos:
            print("⚠️ Nenhum arquivo recebido para processamento.")
            return None

        resultado = {}  # Dicionário final com todas as abas concatenadas

        for caminho in lista_arquivos:
            print(f"\n📂 Lendo arquivo: {caminho}")

            try:
                # Carrega as abas do arquivo atual
                xls = pd.ExcelFile(caminho)
                abas = xls.sheet_names
                print(f"   🔹 Abas encontradas: {abas}")

                for aba in abas:
                    df = pd.read_excel(caminho, sheet_name=aba)

                    # Se a aba já existe no resultado → concatena
                    if aba in resultado:
                        print(f"   ➕ Concatenando aba existente: {aba}")
                        resultado[aba] = pd.concat([resultado[aba], df], ignore_index=True)
                    else:
                        print(f"   🆕 Criando nova aba: {aba}")
                        resultado[aba] = df

            except Exception as e:
                print(f"❌ Erro ao ler o arquivo {caminho}: {str(e)}")

        print("\n✅ Concatenação concluída com sucesso!")
        print(f"Abas finais: {list(resultado.keys())}")

        return resultado



import pandas as pd

# Carrega as duas planilhas
planilha1 = pd.read_excel('Planilha Julho 09.10.xlsx', sheet_name='BASE')
planilha2 = pd.read_excel('Planilha Julho 09.10 1.xlsx', sheet_name='BASE')

# Filtra os usuários com status "lida" em cada planilha
usuarios_lida_1 = set(planilha1.loc[planilha1['Status'] == 'Lida', 'USUARIO'])
usuarios_lida_2 = set(planilha2.loc[planilha2['Status'] == 'Lida', 'USUARIO'])

# Encontra os usuários que estão com "lida" em uma planilha e não na outra
diferenca_1_para_2 = usuarios_lida_1 - usuarios_lida_2
diferenca_2_para_1 = usuarios_lida_2 - usuarios_lida_1

# Salva os resultados em arquivos de texto
with open('lida_na_1_nao_na_2.txt', 'w', encoding='utf-8') as f1:
    for usuario in diferenca_1_para_2:
        f1.write(f"{usuario}\n")

with open('lida_na_2_nao_na_1.txt', 'w', encoding='utf-8') as f2:
    for usuario in diferenca_2_para_1:
        f2.write(f"{usuario}\n")

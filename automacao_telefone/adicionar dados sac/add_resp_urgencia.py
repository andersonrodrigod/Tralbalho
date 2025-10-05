import pandas as pd
import warnings
warnings.simplefilter("ignore", UserWarning)

origem = "detalhamento_de_pesquisa_julho_urgencia.xlsx"
destino = "Planilha Julho 7 01_10.xlsx"

planilhas_origem = pd.read_excel(origem, sheet_name=None)
planilhas_destino = pd.read_excel(destino, sheet_name=None)

abas = {
    "p1": "1_Como_avalia_o_momento_da_s",
    "comen p1": "Nos_relate_o_motivo_da_sua_i",
    "p2": "2_Como_avalia_o_atendimento_",
    "comen p2": "Nos_relate_o_motivo_da_sua_i 1",
    "p3": "3_Como_avalia_a_equipe_de_en",
    "comen p3": "Nos_relate_a_sua_insatisfaca",
    "p4": "4_Como_avalia_a_equipe_medic",
    "comen p4": "Nos_relate_a_sua_insatisfaca 1",
    "p5": "5_Como_avalia_os_servicos_de",
    "comen p5": "Nos_relate_o_motivo_da_sua_i 2"
}


faltando = {chave: valor for chave, valor in abas.items() if valor not in planilhas_origem}

if not faltando:
    print("✅ Todas as abas estão presentes!")
    resultado = True
else:
    print("⚠️ Abas faltando:", faltando)

abas_para_salvar = {}

for aba_destino, aba_origem in abas.items():
    if aba_origem in planilhas_origem:
        df_origem = planilhas_origem[aba_origem]

        if aba_destino in planilhas_destino:
            # concatena os dados do destino com os novos da origem
            df_final = pd.concat([planilhas_destino[aba_destino], df_origem], ignore_index=True)
            print(f"✅ Aba '{aba_destino}': {len(df_origem)} linhas adicionadas, total agora {len(df_final)}")
        else:
            # se a aba não existe no destino, cria nova
            df_final = df_origem
            print(f"✅ Aba '{aba_destino}' criada com {len(df_final)} linhas")

        # atualiza no dicionário destin
        abas_para_salvar[aba_destino] = df_final
        
    else:
        print(f"⚠️ Aba de origem '{aba_origem}' não encontrada no arquivo {origem}")


for aba_existente, df_existente in planilhas_destino.items():
    if aba_existente not in abas_para_salvar:
        abas_para_salvar[aba_existente] = df_existente


with pd.ExcelWriter(destino, engine="openpyxl", mode="w") as writer:
    for aba, df in abas_para_salvar.items():
        df.to_excel(writer,sheet_name=aba, index=False)

















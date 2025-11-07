import pandas as pd
import warnings

warnings.simplefilter("ignore", UserWarning)

# arquivos
origem = "detalhamento_de_pesquisa_eletivo_limpo_16.xlsx"
destino = "Planilha agosto atualizada 16.10.xlsx"

# mapeamento origem → destino (somente abas que serão atualizadas)
abas = {
    "p1": "1_Como_avalia_o_momento_da_s",
    "comen p1": "Nos_relate_o_motivo_da_sua_i",
    "p2": "2_Como_avalia_o_atendimento_",
    "comen p2": "Nos_relate_o_motivo_da_sua_i 1",
    "p3": "3_Como_avalia_a_equipe_de_en",
    "comen p3": "Nos_relate_o_motivo_da_sua_i 2",
    "p4": "4_Como_avalia_a_equipe_medic",
    "comen p4": "Nos_relate_a_sua_insatisfaca",
    "p5": "5_Como_avalia_os_servicos_de",
    "comen p5": "Nos_relate_o_motivo_da_sua_i 3",
    "P6": "6_Como_avalia_o_processo_de_",
    "comen p6": "Nos_relate_o_motivo_da_sua_i 4"
}

planilhas_origem = pd.read_excel(origem, sheet_name=list(abas.values()))

# atualiza cada aba necessária no arquivo destino concatenando os dados
with pd.ExcelWriter(destino, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for aba_destino, aba_origem in abas.items():
        df_novo = planilhas_origem[aba_origem]

        # lê dados existentes no destino (se existir)
        try:
            df_existente = pd.read_excel(destino, sheet_name=aba_destino)
            df_concat = pd.concat([df_existente, df_novo], ignore_index=True)
        except ValueError:  # aba ainda não existe no destino
            df_concat = df_novo

        # salva a aba concatenada
        df_concat.to_excel(writer, sheet_name=aba_destino, index=False)
        print(f"⚡ Aba '{aba_destino}' atualizada ({len(df_concat)} linhas no total).")

print("\n💾 Atualização concluída!")
print("✅ A aba 'BASE' e outras que não foram mencionadas permaneceram intactas com todas as fórmulas.")

import pandas as pd
import warnings

warnings.simplefilter("ignore", UserWarning)

# --- 1. Carregar a planilha ---
arquivo = "planilhas/relatorio_analitico_de_agendamentos_status.xlsx"
df = pd.read_excel(arquivo)

# --- 2. Remover linhas com "Junho" ou "Agosto" na coluna Contato ---
df = df[~df["Contato"].str.contains("junho|julho", case=False, na=False)]

# --- 3. Atualizar coluna Status quando Respondido = "Sim" ---
df.loc[df["Respondido"].str.lower() == "sim", "Status"] = "Lida"

# --- 4. Limpar dados de colunas específicas (sem apagar as colunas) ---
colunas_limpar = [
    "Conta",
    "Mensagem",
    "Categoria",
    "Template",
    "Protocolo",
    "Agendamento",
    "Data agendamento",
    "agendamento",
    "Campanha",
    "Agente"
    "Status agendamento"
]

# Limpa apenas as colunas que realmente existem
colunas_existentes = [c for c in colunas_limpar if c in df.columns]
df[colunas_existentes] = ""

# --- 5. Salvar resultado final ---
saida = "planilha_tratada_status.xlsx"
df.to_excel(saida, index=False)

print("Processo concluído!")
print(f"Linhas com 'Junho' ou 'Agosto' removidas.")
print(f"Status atualizado para 'Lida' onde Respondido = 'Sim'.")
print(f"Colunas limpas: {colunas_existentes}")
print(f"Arquivo salvo como: {saida}")

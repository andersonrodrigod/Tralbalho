import pandas as pd
import pyperclip
import time
import pyautogui as py
from automacao import pegar_telefone, automacao_codigo_inicio, automacao_codigo_next, automacao_codigo_next_sem_dado, aplicar_filtro, verificar_cor_pixel
from checar_dados import ajustar_numero_telefone, is_numero_telefone


def copy_vazio():
    pyperclip.copy("")

def coordenadas_telefone():
    return [
        (79, 545),  # Telefone 1
        (79, 561),  # Telefone 2
        (79, 580),  # Telefone 3
    ]

def automar_fuction(df):
    
    df = pd.read_excel(df, dtype=str)

    time.sleep(2)

    colunas_telefone = ["Telefone 1", "Telefone 2", "Telefone 3", "Telefone 4"]
    df[colunas_telefone] = df[colunas_telefone].fillna("").astype(str).apply(lambda col: col.str.strip())
    df["Status"] = df["Status"].fillna("").astype(str).str.strip()
    

    
    for i, row in df.iterrows():
        codigo = str(row["Codigo"]).strip()
        automacao_codigo_inicio(codigo)
        #print(codigo)
        copy_vazio()
        telefone_adicionado = False
        consecutivos_invalidos = 0
        repeticoes_telefone = 0
        filtro_aplicado = True

        time.sleep(0.7)
        py.hotkey("ctrl", "c")
        time.sleep(0.5)
        conteudo_copy = pyperclip.paste().strip()

        if not conteudo_copy:
            #print(f"Código {codigo}: sem conteúdo no Ctrl+C. Pulando para próxima linha.")
            df.at[i, "Status"] = "BASE INCORRETA"
            #print(f"Código {codigo}: sem conteúdo no Ctrl+C. Pulando para próxima linha.")
            automacao_codigo_next_sem_dado() 
            continue 
        
        numero_preto = None
        for j in range(6):
            if j == 0 and filtro_aplicado:
                aplicar_filtro()

            if j > 2:
                py.click(361, 581)
                time.sleep(0.5)
                x, y = coordenadas_telefone()[2] 
            else:
                x, y = coordenadas_telefone()[j]
                    
            copy_vazio()
            py.click(x, y)
            time.sleep(0.5)
            telefone = pegar_telefone()
            telefone = ajustar_numero_telefone(telefone)
            resultado = is_numero_telefone(telefone)
            telefones_existentes = [row[col] for col in colunas_telefone]
            py.click(x, y)
            cor = verificar_cor_pixel()


            if cor == "PRETO" and resultado == "NOVO":
                numero_preto = telefone
                print("Número preto armazenado.")

            elif cor == "VERDE" and resultado == "NOVO" and telefone not in telefones_existentes:
                print("Número diferente, levando para o Excel.")
                for col in ["Telefone 2", "Telefone 3", "Telefone 4"]:
                    if row[col] == "":
                        df.at[i, col] = telefone
                        print(f"Código {codigo}: número adicionado em {col}")
                        df.at[i, "Status"] = "NOVO CONTATO"
                        telefone_adicionado = True
                        break
                if telefone_adicionado:
                    break

            elif telefone in telefones_existentes:
                repeticoes_telefone += 1
                print(f"Número igual ({repeticoes_telefone}x), tentando novamente...")

                if repeticoes_telefone >= 4:
                    filtro_aplicado = False
                    print("Número repetido 4 vezes. Encerrando.")
                    break

            else:
                print("Número inválido")
                consecutivos_invalidos += 1

                if consecutivos_invalidos >= 1:
                    print("Dois números inválidos consecutivos. Encerrando.")
                    py.press("f4")
                    filtro_aplicado = False
                    break

        if not telefone_adicionado and not filtro_aplicado:
            py.press("f8")
            numero_preto_segundo_loop = None
            for j in range(8):
                if j > 2:
                    py.click(361, 581)
                    time.sleep(0.5)
                    x, y = coordenadas_telefone()[2]
                else:
                    x, y = coordenadas_telefone()[j]

                copy_vazio()
                py.click(x, y)
                time.sleep(0.5)
                telefone = pegar_telefone()
                telefone = ajustar_numero_telefone(telefone)
                resultado = is_numero_telefone(telefone)

                # mesma lógica de salvamento aqui

                py.click(x, y)
                telefones_existentes = [row[col] for col in colunas_telefone]

                cor = verificar_cor_pixel()

                if cor == "PRETO" and resultado == "NOVO":
                    numero_preto_segundo_loop = telefone
                    print("Número preto armazenado.")

                if cor == "VERDE" and resultado == "NOVO" and telefone not in telefones_existentes:
                    print("Número diferente, levando para o Excel.")
                    for col in ["Telefone 2", "Telefone 3", "Telefone 4"]:
                        if row[col] == "":
                            df.at[i, col] = telefone
                            print(f"Código {codigo}: número adicionado em {col}")
                            df.at[i, "Status"] = "NOVO CONTATO"
                            telefone_adicionado = True
                            break
                    if telefone_adicionado:
                        break
                elif numero_preto and numero_preto not in telefones_existentes:
                    print("Número diferente, levando para o Excel.")
                    for col in ["Telefone 2", "Telefone 3", "Telefone 4"]:
                        if row[col] == "":
                            df.at[i, col] = telefone
                            print(f"Código {codigo}: número adicionado em {col}")
                            df.at[i, "Status"] = "NOVO CONTATO"
                            telefone_adicionado = True
                            break
                    if telefone_adicionado:
                        break

                elif numero_preto_segundo_loop and numero_preto_segundo_loop not in telefones_existentes:
                    print("Número diferente, levando para o Excel.")
                    for col in ["Telefone 2", "Telefone 3", "Telefone 4"]:
                        if row[col] == "":
                            df.at[i, col] = telefone
                            print(f"Código {codigo}: número adicionado em {col}")
                            df.at[i, "Status"] = "NOVO CONTATO"
                            telefone_adicionado = True
                            break
                    if telefone_adicionado:
                        break

                elif telefone in telefones_existentes:
                    repeticoes_telefone += 1
                    print(f"Número igual ({repeticoes_telefone}x), tentando novamente...")
                    
                    if df.at[i, "Status"] != "NOVO CONTATO":
                        df.at[i, "Status"] = "MESMO CONTATO"

                    if repeticoes_telefone >= 4:
                        print("Número repetido 4 vezes. Encerrando.")
                        break
                else:
                    print("Número inválido")
                    consecutivos_invalidos += 1

                    if consecutivos_invalidos >= 2:
                        print("Dois números inválidos consecutivos. Encerrando.")
                        
                        if df.at[i, "Status"] not in ["NOVO CONTATO", "MESMO CONTATO"]:
                            df.at[i, "Status"] = "SEM CONTATO"
                        break 
                    

        telefones_atualizados = [df.at[i, c] for c in colunas_telefone]
        #print(f"Código: {codigo} | Tel1: {telefones_atualizados[0]} | Tel2: {telefones_atualizados[1]} | Tel3: {telefones_atualizados[2]}")

        # Salvar checkpoint a cada 1 linhas
        if i % 2 == 0:
            df.to_excel("dados_clinipan.xlsx", index=False)
            #print(f"Checkpoint salvo na linha {i}")


        automacao_codigo_next()  
        
    
    df.to_excel("dados_clinipan.xlsx", index=False)
    #print("Salvamento final concluído.")


dados = "dados_clinipan.xlsx"

automar_fuction(dados)



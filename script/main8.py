import pandas as pd
import os
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk

# Função para encontrar o melhor salto
def melhor_salto(val1, val2):
    return max(val1, val2)

# Função para processar os arquivos e consolidar os dados
def processar_arquivos(diretorio):
    dataframes = []
    print(f"Processando arquivos no diretório: {diretorio}")

    for arquivo in os.listdir(diretorio):
        if arquivo.startswith("JumpSystem-Avaliação") and arquivo.endswith(".xlsx"):
            caminho = os.path.join(diretorio, arquivo)
            print(f"Lendo arquivo: {caminho}")
            df = pd.read_excel(caminho, header=None, skiprows=3)

            # Verificação das linhas e colunas
            if df.shape[0] > 9 and df.shape[1] > 10:
                nome = pd.read_excel(caminho, header=None).iat[2, 1]
                sexo = pd.read_excel(caminho, header=None).iat[2, 10]
                idade = pd.read_excel(caminho, header=None).iat[2, 9]
                
                dados = {
                    'Nome': nome,
                    'Sexo': sexo,
                    'Idade': idade,
                    'Potência Squat Jump': melhor_salto(df.iat[4, 4], df.iat[5, 4]) if df.shape[0] > 5 and df.shape[1] > 4 else None,
                    'Altura Squat Jump': melhor_salto(df.iat[4, 3], df.iat[5, 3]) if df.shape[0] > 5 and df.shape[1] > 3 else None,
                    'Potência Contramov': melhor_salto(df.iat[6, 4], df.iat[7, 4]) if df.shape[0] > 7 and df.shape[1] > 4 else None,
                    'Altura Contramov': melhor_salto(df.iat[6, 3], df.iat[7, 3]) if df.shape[0] > 7 and df.shape[1] > 3 else None,
                    'Potência CMJ MMSS': melhor_salto(df.iat[8, 4], df.iat[9, 4]) if df.shape[0] > 9 and df.shape[1] > 4 else None,
                    'Altura CMJ MMSS': melhor_salto(df.iat[8, 3], df.iat[9, 3]) if df.shape[0] > 9 and df.shape[1] > 3 else None,
                }
                print(f"Dados extraídos: {dados}")
                dataframes.append(pd.DataFrame(dados, index=[0]))
            else:
                print(f"Arquivo {arquivo} ignorado: formato inesperado")

    if dataframes:
        df_consolidado = pd.concat(dataframes, ignore_index=True)
        df_consolidado.to_excel('Banco_de_dados_GERAL.xlsx', index=False)
        print("Arquivo consolidado gerado com sucesso!")
    else:
        print("Nenhum dado válido encontrado para consolidar.")

# Função para abrir o diálogo de seleção de diretório
def escolher_diretorio():
    diretorio = filedialog.askdirectory()
    if diretorio:
        processar_arquivos(diretorio)

# Configuração da interface gráfica
ctk.set_appearance_mode("dark")  # Modo escuro
ctk.set_default_color_theme("dark-blue")  # Tema azul escuro

app = ctk.CTk()  # Criação da janela principal
app.geometry("400x200")
app.title("Consolidador de Arquivos JumpSystem")

label = ctk.CTkLabel(app, text="Selecione o diretório dos arquivos JumpSystem-Avaliação", pady=20)
label.pack()

botao = ctk.CTkButton(app, text="Escolher Diretório", command=escolher_diretorio)
botao.pack(pady=20)

app.mainloop()

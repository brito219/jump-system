import pandas as pd
import os
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk

def maximum(val1, val2):
    return max(val1, val2)

def select(arquivos):
    dataframes = []
    print(f"Processando arquivos selecionados...")

    for path in arquivos:
        print(f"Lendo arquivo: {path}")
        df = pd.read_excel(path, header=None)

        if df.shape[0] > 9 and df.shape[1] > 13:
            nome = pd.read_excel(path, header=None).iat[2, 2]
            sexo = pd.read_excel(path, header=None).iat[2, 11]
            idade = pd.read_excel(path, header=None).iat[2, 10]
            
            dados = {
                'Nome': nome,
                'Sexo': sexo,
                'Idade': idade,
                'Potência Squat Jump': maximum(df.iat[7, 4], df.iat[8, 4]),
                'Altura Squat Jump': maximum(df.iat[7, 3], df.iat[8, 3]),
                'Potência Contramov': maximum(df.iat[9, 4], df.iat[10, 4]),
                'Altura Contramov': maximum(df.iat[9, 3], df.iat[10, 3]),
                'Potência CMJ MMSS': maximum(df.iat[11, 4], df.iat[12, 4]),
                'Altura CMJ MMSS': maximum(df.iat[11, 3], df.iat[12, 3]),
            }
            print(f"Dados extraídos: {dados}")
            dataframes.append(pd.DataFrame(dados, index=[0]))
        else:
            print(f"Arquivo {path} ignorado: formato inesperado")

    if dataframes:
        df_consolidado = pd.concat(dataframes, ignore_index=True)
        df_consolidado.to_excel('Banco_de_dados_GERAL.xlsx', index=False)
        print("Arquivo consolidado gerado com sucesso!")
    else:
        print("Nenhum dado válido encontrado para consolidar.")

def escolher_arquivos():
    arquivos = filedialog.askopenfilenames(title="Selecione os arquivos de avaliação", 
                                           filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivos:
        select(arquivos)

# Configuração da interface gráfica
ctk.set_appearance_mode("dark")  # Modo escuro
ctk.set_default_color_theme("dark-blue")  # Tema azul escuro

app = ctk.CTk()  # Criação da janela principal
app.geometry("400x200")
app.title("Consolidador de Arquivos JumpSystem")

label = ctk.CTkLabel(app, text="Selecione os arquivos JumpSystem-Avaliação", pady=20)
label.pack()

botao = ctk.CTkButton(app, text="Escolher Arquivos", command=escolher_arquivos)
botao.pack(pady=20)

app.mainloop()

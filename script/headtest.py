import pandas as pd

# Carregar o arquivo de exemplo
caminho_arquivo = 'C:/Users/Brito/Documents/Estudos/Projeto de Extensão Joel/teste/JumpSystem-Avaliação3.xlsx'
df_exemplo = pd.read_excel(caminho_arquivo, header=None)

# Exibir as primeiras linhas do DataFrame
print(df_exemplo.head(15))

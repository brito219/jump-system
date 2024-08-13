import pandas as pd

# Carregar o arquivo de exemplo
caminho_arquivo = 'diretorio_avaliacoes\JumpSystem-Avaliação837.xlsx'
df_exemplo = pd.read_excel(caminho_arquivo, header=None)

# Exibir as primeiras linhas do DataFrame
print(df_exemplo.head(15))

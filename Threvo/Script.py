import pandas as pd
import glob
import os

#Diretório contendo os arquivos Excel
diretorio = r'C:\Users\Lucas\OneDrive\Documentos\GitHub\Projects\Python\Threvo\Diretorio'

#Lista para armazenar os caminhos dos arquivos Excel
arquivos_excel = glob.glob(os.path.join(diretorio, '*.xlsx'))

#Verificação se há arquivos Excel no diretório
if not arquivos_excel:
    print("Nenhum arquivo Excel encontrado no diretório especificado.")
else:
    #Inicializa o DataFrame que irá armazenar os dados
    dados_historicos = pd.DataFrame()

    #Loop através de cada arquivo Excel e concatena seus dados ao DataFrame
    for arquivo_excel in arquivos_excel:
        try:
            tabela = pd.read_excel(arquivo_excel)
            dados_historicos = pd.concat([dados_historicos, tabela], ignore_index=True)
            print(f"Arquivo '{os.path.basename(arquivo_excel)}' lido e adicionado.")
        except Exception as e:
            print(f"Erro ao ler o arquivo '{os.path.basename(arquivo_excel)}': {e}")

    #Salvar o DataFrame resultante em um único arquivo Excel
    nome_arquivo_saida = os.path.join(diretorio, 'dados_compilados.xlsx')
    dados_historicos.to_excel(nome_arquivo_saida, index=False)
    print(f"Dados compilados salvos em '{nome_arquivo_saida}'.")

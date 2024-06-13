import os
import pandas as pd

# Função para ler todos os DataFrames de uma pasta e adicionar o nome do arquivo como uma coluna
def read_all_dataframes_from_folder(folder_path):
    dataframes = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(file_path)
            df['filename'] = filename  # Adicionar o nome do arquivo como uma nova coluna
            dataframes.append(df)
    return dataframes

# Função para unificar todos os DataFrames em um único DataFrame
def unify_dataframes(dataframes):
    unified_df = pd.concat(dataframes, ignore_index=True)
    return unified_df

# Função para salvar o DataFrame unificado em um arquivo Excel
def save_dataframe_to_folder(dataframe, output_folder, output_filename):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    output_path = os.path.join(output_folder, output_filename)
    dataframe.to_excel(output_path, index=False)

# Ler o arquivo de relação de arquivos
relacao_arquivos = pd.read_csv("GIKFollow.csv")

for _, row in relacao_arquivos.iterrows():
    input_folder = row['caminho_entrada']
    output_folder = row['caminho_saida']
    ano = row['ano']

    # Ler todos os DataFrames da pasta de entrada
    dataframes = read_all_dataframes_from_folder(input_folder)

    # Unificar todos os DataFrames
    unified_dataframe = unify_dataframes(dataframes)

    # Definir o nome do arquivo de saída
    output_filename = f"KPI's {ano}.xlsx"

    # Salvar o DataFrame unificado na pasta de saída
    save_dataframe_to_folder(unified_dataframe, output_folder, output_filename)

    print(f"DataFrames unificados e salvos com sucesso em {output_folder} como {output_filename}!")

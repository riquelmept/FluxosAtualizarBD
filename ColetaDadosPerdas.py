import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk

def processar_aba(df, vars_id, prefixo_col_data):
    """
    Converte um DataFrame do formato largo para o formato longo, processando e formatando colunas de data.

    Parâmetros:
    df (DataFrame): O DataFrame do pandas a ser processado.
    vars_id (list): Lista de colunas a serem usadas como variáveis identificadoras.
    prefixo_col_data (str): Uma string que as colunas contendo datas começam.

    Retorna:
    DataFrame: O DataFrame do pandas processado no formato longo com as colunas de data formatadas.
    Função Específca para o trabalho com a planilha das 5 plantas da LDC.
    """
    # Encontrar as colunas de data pelo prefixo_col_data
    colunas_data = [col for col in df.columns if str(col).startswith(prefixo_col_data)]
    
    # Transformando o dataframe do formato largo para o longo (Unipivot)
    df_transformado = df.melt(id_vars=vars_id, value_vars=colunas_data,
                              var_name="Dia", value_name="Duração")
    
    # Convertendo a coluna 'Data'
    df_transformado['Dia'] = pd.to_datetime(df_transformado['Dia'], errors='coerce').dt.date

    # Descartar qualquer linha que seja luna
    df_transformado = df_transformado.dropna(subset=['Dia'])
    
    return df_transformado

def processar_todas_abas(caminho_arquivo, vars_id, prefixo_col_data):
    """
    Lê um arquivo Excel, processa cada aba e concatena em um único DataFrame.

    Parâmetros:
    caminho_arquivo (str): O caminho para o arquivo Excel.
    vars_id (list): Lista de colunas a serem usadas como variáveis identificadoras.
    prefixo_col_data (str): Uma string que as colunas contendo datas começam.

    Retorna:
    DataFrame: O DataFrame concatenado contendo dados de todas as abas.
    Função Específca para o trabalho com a planilha das 5 plantas da LDC.

    """
    # Carregar o arquivo Excel
    xls = pd.ExcelFile(caminho_arquivo)

    # Ler todas as abas para um dicionário de DataFrames
    abas_dict = pd.read_excel(xls, sheet_name=None)

    # Inicializar um DataFrame vazio para manter os dados unificados
    df_unificado = pd.DataFrame()

    # Processar cada aba e concatenar em um DataFrame unificado
    for nome_aba, df in abas_dict.items():
        df_processado = processar_aba(df, vars_id, prefixo_col_data)
        df_unificado = pd.concat([df_unificado, df_processado], ignore_index=True)
        
    return df_unificado

def agrupar_falhas(df):
    """
    Agrupa as falhas por planta, categoria e data, e conta o número de falhas.

    Parâmetros:
    df (DataFrame): O DataFrame com os dados das falhas.

    Retorna:
    DataFrame: O DataFrame agrupado com a contagem de falhas.
    """
    # Excluir a coluna 'Setor'
    df = df.drop(columns=['Departamento','Sistema', 'Ativo'])

    df = df[df['Duração'] != 0]
    
    # Agrupar por planta, categoria e data, e contar o número de falhas
    df_agrupado = df.groupby(['Und', 'Caráter', 'Dia']).size().reset_index(name='Contagem_Falhas')
    
    return df_agrupado

tabela = pd.read_csv("arquivos.csv")
print(tabela)
arq_geral = pd.DataFrame()

for linha in tabela.index:
    if tabela.loc[linha, 'planta'] != 'Geral':
        caminho_arquivo = tabela.loc[linha, 'caminho_arq']
        vars_id = ['Planta', 'Setor', 'Categoria']
        prefixo_colunas_data = '2024'
        df_unificado = processar_todas_abas(caminho_arquivo, vars_id, prefixo_colunas_data)
        caminho_arquivo_saida = tabela.loc[linha, 'ttd_arq_caminho']
        df_unificado.to_excel(caminho_arquivo_saida, index=False)
        print(df_unificado.head())
    else:
        df_perdas = pd.read_excel(tabela.loc[linha, 'caminho_arq'])
        df_perdas = agrupar_falhas(df_perdas)
        df_perdas.head()
        caminho_alternativo = tabela.loc[linha, 'ttd_arq_caminho']
        df_perdas.to_excel(caminho_alternativo, index = False)
        df_perdas.to_excel('BD Quantidade de Paradas.xlsx')

# Criar a janela principal
janela = tk.Tk()
janela.title("Atualização Finalizada!")
# Criar o rótulo com a mensagem
mensagem = tk.Label(janela, text="Atualização Finalizada!", font=("Arial", 14))
mensagem.pack(padx=10, pady=10)

# Executar o loop principal da interface gráfica
janela.mainloop()

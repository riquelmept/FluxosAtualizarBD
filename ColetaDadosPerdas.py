import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

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
                              var_name="Data", value_name="Valor")
    
    # Convertendo a coluna 'Data'
    df_transformado['Data'] = pd.to_datetime(df_transformado['Data'], errors='coerce').dt.date

    # Descartar qualquer linha que seja luna
    df_transformado = df_transformado.dropna(subset=['Data'])
    
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
    df = df.drop(columns=['Setor'])

    df = df[df['Valor'] != 0]
    
    # Agrupar por planta, categoria e data, e contar o número de falhas
    df_agrupado = df.groupby(['Planta', 'Categoria', 'Data']).size().reset_index(name='Contagem_Falhas')
    
    return df_agrupado

tabela = pd.read_csv("arquivos.csv")
print(tabela)
arq_geral = pd.DataFrame()

for linha in tabela.index:
    caminho_arquivo = tabela.loc[linha, "caminho_arq"]
    vars_id = ["Planta", "Setor", "Categoria"]
    prefixo_colunas_data = "2024"
    df_unificado = processar_todas_abas(caminho_arquivo, vars_id, prefixo_colunas_data)
    caminho_arquivo_saida = tabela.loc[linha, "ttd_arq_caminho"]
    df_unificado.to_excel(caminho_arquivo_saida, index=False)
    print(df_unificado.head())

    arq_geral = pd.concat([arq_geral, df_unificado], ignore_index= True)

arq_geral = agrupar_falhas(arq_geral)
arq_geral.to_excel('BD Quantidade de Paradas.xlsx')

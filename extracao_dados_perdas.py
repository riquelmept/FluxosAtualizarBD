import pandas as pd

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

# Exemplo de uso:
# Defina o caminho para o arquivo excel que vai processar(Caminho das pastas das Plantas)
caminho_arquivo = input(str('Insira aqui o caminho do arquivo que vai ser tratado'))
# Defina as colunas identificadoras e o prefixo das suas colunas de data
vars_id = ["Planta", "Setor", "Categoria"]
prefixo_colunas_data = "2024"  # Use o ano apropriado ou prefixo para identificar suas colunas de data

# Processar o arquivo e obter o DataFrame unificado
df_unificado = processar_todas_abas(caminho_arquivo, vars_id, prefixo_colunas_data)

# Salvar o DataFrame processado em um novo arquivo Excel (Caminho onde será salvo o arquivo tratado)
caminho_arquivo_saida = input(str('Insira aqui o caminho onde salvar o novo arquivo'))
df_unificado.to_excel(caminho_arquivo_saida, index=False)

# Exibir as primeiras linhas do DataFrame unificado
print(df_unificado.head())

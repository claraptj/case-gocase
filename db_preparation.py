import os
import glob
import pandas as pd
import unicodedata

# Função para normalizar strings e remover acentos
def remover_acentos(texto):
    if isinstance(texto, str):
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto

# 1. Obtém dados de planilhas Excel
matching_files = glob.glob("BASE CORPORATIVA*.xlsx") # Busca por arquivos aplicáveis, dentro da pasta atual
matching_files.sort(key=os.path.getmtime, reverse=True) # Ordena por data de modificação, do mais recente para o mais antigo
if matching_files:
    # Caso a lista de arquivos não esteja vazia
    file_name = matching_files[0] # Seleciona o arquivo mais recente
    print(f"Obtendo dados do arquivo '{file_name}'...")

    try:
        working_sheet="Página1"
        df = pd.read_excel(file_name, sheet_name=working_sheet) # Lê somente a aba com nome pré-definido
        if df.empty:
            # Caso a aba esteja vazia, levanta um erro
            raise ValueError(f"A aba '{working_sheet}' não contém dados.")
        else:
            # Caso a aba tenha dados, imprime o número de linhas carregadas
            print(f"\tDados de '{working_sheet}' obtidos com sucesso. \n\tTotal de {len(df)} linhas carregadas.")
            # # print(df) # Imprime o dataframe para validação
    except Exception as e:
        print(f"\tErro ao ler a planilha: {e}")
else:
    raise FileNotFoundError("Não foram encontradas planilhas Excel com o padrão 'BASE CORPORATIVAS*.xlsx'.")

# 2. Renomeia colunas para padronização
df.columns = [
    remover_acentos(col).strip().lower().replace(' ', '_')
    for col in df.columns
]
# # print(df) # Imprime o dataframe para validação

# 3. Conversões e tratamento
df[['nome_do_negocio', 'qtd_do_negocio']] = df['nome_do_negocio'].str.extract(r'^(.*)\s+\(Qtd:\s*(\d+)\)$') # Extrai nome e quantidade do negócio para colunas distintas

numeric_columns = ['qtd_do_negocio', 'valor', 'numero_de_atividades_de_vendas']
df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce') # Formata colunas como número

datetime_columns = ['data_de_fechamento', 'data_da_ultima_modificacao']
df[datetime_columns] = df[datetime_columns].apply(lambda col: pd.to_datetime(col, format='%Y-%m-%d %H:%M:%S', errors='coerce'))  # Formata colunas como data/hora

df = df.sort_values(by='data_da_ultima_modificacao') # Ordena o dataframe pela data da última modificação
df = df.sort_values(by='data_da_ultima_modificacao').reset_index(drop=True) # Reseta o índice do dataframe após ordenação

# # print(df[numeric_columns]) # Imprime o dataframe para validação
# # print(df.dtypes) # Imprime tipos de dados do dataframe para validação

# 4. Exporta base tratada para Power BI
subfolder_name = 'bases_para_analise'
new_file_name = 'base_atualizada'
if subfolder_name:
    # Caso seja definido um nome para a subpasta, o inclui no caminho do arquivo
    os.makedirs(subfolder_name, exist_ok=True)  # Cria a subpasta caso ela não exista
    file_path = os.path.join(subfolder_name, f'{new_file_name}.csv')
else:
    file_path = f'{new_file_name}.csv'

df.index.name = 'codigo_da_proposta'  # Define o nome da coluna de índice
df.to_csv(file_path, index=True, encoding='utf-8') # Exporta o dataframe para CSV
print(f"Arquivo exportado com sucesso: '{file_path}'.")

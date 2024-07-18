import pandas as pd

# Ajustar as configurações de exibição do Pandas
pd.set_option('display.max_columns', None)  # Mostrar todas as colunas
pd.set_option('display.max_rows', None)     # Mostrar todas as linhas (cuidado com DataFrames grandes)
pd.set_option('display.max_colwidth', None) # Mostrar todo o conteúdo das colunas

# Carregar o arquivo Excel
file_path = 'C:\PyCharmProjects\AppCA\My Estimate - 2 configs.xlsx'
excel_data = pd.ExcelFile(file_path)

# Carregar os dados da planilha 'PAAS', ignorando as três primeiras linhas e usando a linha correta como cabeçalho
data = pd.read_excel(file_path, sheet_name='PAAS', skiprows=4)

# Remover linhas onde 'Part' e 'Description' são ambas NaN
cleaned_data = data.dropna(subset=['Part', 'Description'], how='all')

# Substituir valores vazios por zero nas colunas numéricas antes da conversão
cleaned_data[['Part Qty', 'Instance Qty', 'Usage Qty', 'Unit Price', 'Monthly Cost']] = cleaned_data[
    ['Part Qty', 'Instance Qty', 'Usage Qty', 'Unit Price', 'Monthly Cost']].fillna(0)

# # Converter colunas numéricas para os tipos de dados apropriados
# numeric_columns = ['Part Qty', 'Instance Qty', 'Usage Qty', 'Unit Price', 'Monthly Cost']
# for column in numeric_columns:
#     # Remover caracteres não numéricos de 'Monthly Cost' (e.g., '$') e converter para float
#     if column == 'Monthly Cost':
#         cleaned_data[column] = cleaned_data[column].replace('[\$,]', '', regex=True).astype(float)
#     else:
#         cleaned_data[column] = pd.to_numeric(cleaned_data[column], errors='coerce')

# Criar uma nova coluna para o nome do serviço pai
cleaned_data['Parent Service'] = None

# Inicializar uma variável para armazenar o nome do serviço pai atual
current_parent_service = None

# Percorrer as linhas do dataframe
for index, row in cleaned_data.iterrows():
    description = row['Description']
    part = row['Part']

    # Verificar se a linha atual é um serviço pai
    if pd.isna(part):
        current_parent_service = description
        cleaned_data.at[index, 'Parent Service'] = current_parent_service
    else:
        # Atribuir o serviço pai atual à linha do serviço filho
        cleaned_data.at[index, 'Parent Service'] = current_parent_service

# Exibir os dados resultantes
print(cleaned_data.head(20))

# Salvar os dados limpos em um novo arquivo Excel
cleaned_data.to_excel('C:\PyCharmProjects\AppCA\Cleaned_Estimate.xlsx', index=False)















# import pandas as pd
#
# # Ajustar as configurações de exibição do Pandas
# pd.set_option('display.max_columns', None)  # Mostrar todas as colunas
# pd.set_option('display.max_rows', None)     # Mostrar todas as linhas (cuidado com DataFrames grandes)
# pd.set_option('display.max_colwidth', None) # Mostrar todo o conteúdo das colunas
#
# # Carregar o arquivo Excel
# file_path = 'C:\PyCharmProjects\AppCA\My Estimate - 2 configs.xlsx'
# df = pd.read_excel(file_path, sheet_name='PAAS', skiprows=4)
#
# # Limpar dados vazios: remover linhas completamente vazias
# df_cleaned = df.dropna(how='all')
#
# # Separar os serviços
# base_database_service = df_cleaned[df_cleaned['Description'].str.contains('Base Database Service - Virtual Machine', na=False)]
# virtual_machine_service = df_cleaned[df_cleaned['Description'].str.contains('Virtual Machine', na=False) &
#                                      ~df_cleaned['Description'].str.contains('Base Database Service - Virtual Machine', na=False)]
#
# # Exibir os DataFrames separados
# print("Base Database Service - Virtual Machine")
# print(base_database_service)
#
# print("\nVirtual Machine")
# print(virtual_machine_service)



















# import pandas as pd
#
# # Ajustar as configurações de exibição do Pandas
# pd.set_option('display.max_columns', None)  # Mostrar todas as colunas
# pd.set_option('display.max_rows', None)     # Mostrar todas as linhas (cuidado com DataFrames grandes)
# pd.set_option('display.max_colwidth', None) # Mostrar todo o conteúdo das colunas
#
#
# # Caminho para o arquivo Excel
# file_path = 'C:\PyCharmProjects\AppCA\My Estimate - 2 configs.xlsx'
#
# # Carregar os dados a partir da linha correta e definir o cabeçalho
# df = pd.read_excel(file_path, sheet_name='PAAS', skiprows=4)
#
# # Limpar dados vazios: remover linhas completamente vazias
# df_cleaned = df.dropna(how='all')

#print(df_cleaned)

# # Preencher células vazias em colunas específicas com valores padrão
# df_cleaned['Part Qty'].fillna(0, inplace=True)
# df_cleaned['Instance Qty'].fillna(0, inplace=True)
# df_cleaned['Usage Qty'].fillna(0, inplace=True)
# df_cleaned['Unit Price'].fillna(0.0, inplace=True)
# df_cleaned['Monthly Cost'].fillna('0 US$', inplace=True)
#
# # Converter tipos de dados: converter preços para float após remover símbolos de moeda e espaços
# df_cleaned['Unit Price'] = df_cleaned['Unit Price'].replace('[\$,]', '', regex=True).astype(float)
# df_cleaned['Monthly Cost'] = df_cleaned['Monthly Cost'].replace('[\$,]', '', regex=True).astype(float)
#
# # Formatar valores: por exemplo, formatar a coluna de custos mensais para exibir como moeda
# df_cleaned['Monthly Cost'] = df_cleaned['Monthly Cost'].apply(lambda x: f"${x:,.2f}")
#
# # Salvar em um novo arquivo Excel
# output_file_path = 'C:\PyCharmProjects\AppCA\My Estimate2 - 2 configs.xlsx'
# df_cleaned.to_excel(output_file_path, index=False)
#
# output_file_path

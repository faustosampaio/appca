import pandas as pd

# Carregar o arquivo Excel
file_path = 'Fausto6.xlsx'
xls = pd.ExcelFile(file_path)

# Carregar a planilha em um DataFrame, ignorando as primeiras 4 linhas
df = pd.read_excel(file_path, sheet_name='PAAS', skiprows=4)

# Remover colunas completamente vazias
df.dropna(how='all', axis=1, inplace=True)

# Remover o prefixo 'R$' da coluna 'Monthly Cost'
df['Monthly Cost'] = df['Monthly Cost'].str.replace('R$', '', regex=False).str.strip()

# Inicializar um dicionário vazio para armazenar os serviços
services = {}
current_service = None

# Iterar sobre as linhas do DataFrame
for index, row in df.iterrows():
    # Se a coluna 'Part' for NaN, pode ser um título de serviço
    if not str(row['Part']).startswith('B'):
        current_service = str(row['Description']).strip()
        if current_service != 'Monthly Total':
            services[current_service] = []
    elif current_service:
        # Adicionar os detalhes do serviço ao serviço atual
        services[current_service].append(row.to_dict())

# Remover entradas onde 'current_service' seja 'NaN'
services = {k: v for k, v in services.items() if k and k != 'nan'}

# Exibir os serviços extraídos
for service, details in services.items():
    print(f"Service: {service}")
    for detail in details:
        print(detail)
    print()

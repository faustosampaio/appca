import pandas as pd
import json

# Carregar o arquivo Excel
excel_file_path = 'C:\PyCharmProjects\AppCA\Fausto6.xlsx'
excel_data = pd.read_excel(excel_file_path)

# Estrutura de exemplo baseada no JSON fornecido
json_data = {
    "label": "Fausto6",
    "timeFrame": {
        "months": 1,
        "from": None,
        "to": None
    },
    "currency": "BRL",
    "meta": {
        "exportVersion": 1,
        "dataBuildID": {
            "buildNumber": 1419,
            "instance": "GEN2-PUBLIC-PROD",
            "serviceType": "paas",
            "realm": "public",
            "buildDate": "2024-07-11T09:45:23Z"
        },
        "hash": "c2aceefd15e378862b991f2390571408"
    },
    "configs": []
}

# Função para converter os dados do Excel em JSON no formato desejado
def convert_excel_to_json(excel_data):
    for index, row in excel_data.iterrows():
        service = {
            "label": row['label'],  # Ajuste os nomes das colunas conforme o seu arquivo Excel
            "services": [
                {
                    "id": row['id'],
                    "label": row['service_label'],
                    "utilization": {
                        "instances": row['instances'],
                        "hours": row['hours'],
                        "days": row['days'],
                        "hoursPerMonth": row['hoursPerMonth']
                    },
                    "items": [
                        {
                            "sku": row['sku'],
                            "quantity": row['quantity'],
                            "minQuantity": row['minQuantity'],
                            "inPreset": row['inPreset'],
                            "unitPrice": row['unitPrice'],
                            "monthlyCost": row['monthlyCost']
                        }
                    ],
                    "uiState": {},
                    "presetID": row['presetID'],
                    "isFromShape": row['isFromShape']
                }
            ],
            "addBySearch": False
        }
        json_data["configs"].append(service)

# Converter os dados do Excel
convert_excel_to_json(excel_data)

# Salvar o JSON em um arquivo
json_file_path = 'caminho_para_o_arquivo_json.json'
with open(json_file_path, 'w', encoding='utf-8') as json_file:
    json.dump(json_data, json_file, ensure_ascii=False, indent=4)

print(f"Arquivo JSON salvo em: {json_file_path}")

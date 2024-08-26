import requests
import pandas as pd

# Define os headers da requisição
headers = {
    "identifier": "04.689.450/0001-94",
    "key": "319c0905-9f40-429e-b775-c243dea4e617",
    'User-Agent': 'PostmanRuntime/7.30.0'
}

# Função para fazer a requisição e retornar um DataFrame


def get_data_from_api(url, payload, headers):
    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            data = response.json()
            df = pd.json_normalize(data)
            obj_data = pd.json_normalize(df['Obj'].explode())
            return obj_data
        except ValueError as e:
            print(f"Erro ao decodificar JSON: {e}")
            print("Conteúdo da resposta:")
            print(response.text)
    else:
        print(f"Falha: {response.status_code}")
        print(response.text)
    return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro


# URL dos endpoints
url1 = 'https://www.dimepkairos.com.br/RestServiceApi/People/SearchPeople'
# Exemplo de endpoint diferente
url2 = 'https://www.dimepkairos.com.br/RestServiceApi/Justification/GetJustification'

# Payloads das requisições
payload1 = {"Matricula": 0}
payload2 = {"Code": 0,
            "IdType": 1202,
            "ResponseType": "AS400V1"}  # Modifique conforme necessário

# Faz as requisições e obtém os DataFrames
df1 = get_data_from_api(url1, payload1, headers)
df2 = get_data_from_api(url2, payload2, headers)

# Adiciona um prefixo para diferenciar as colunas dos diferentes DataFrames
df1 = df1.add_prefix('People_')
df2 = df2.add_prefix('Just_')

# Combine os DataFrames
combined_df = pd.concat([df1, df2], axis=1)

# Seleciona apenas as colunas desejadas
selected_columns = [
    'People_Id', 'People_Matricula', 'People_Nome', 'People_Cpf',
    # Adicione aqui as colunas que deseja do segundo DataFrame
    'Just_Id', 'Just_Description',
]

# Filtra o DataFrame combinado para manter apenas as colunas selecionadas
filtered_combined_df = combined_df[selected_columns]

# Exibe o DataFrame filtrado para verificação
print(filtered_combined_df.head())

# Define o caminho completo onde o arquivo Excel será salvo
output_path = r'C:\Users\WI140\Desktop\Nicoly_empresa.xlsx'

# Salva o DataFrame filtrado em um arquivo Excel
filtered_combined_df.to_excel(output_path, index=False)
print(f"Arquivo salvo em: {output_path}")

import requests
import pandas as pd
import openpyxl

# Leitor do arquivo excel (mudar o diretório dependendo do pc)
df = pd.read_excel('Envio_de_Marcacao/marcacaoteste.xlsx')
df['data_hora'] = pd.to_datetime(df['data'].astype(str) + ' ' +
                                 df['hora'].astype(str),
                                 format='mixed')

df['data_hora_iso'] = df['data_hora'].dt.strftime('%d/%m/%Y %H:%M')

# Transformando a leitura das colunas em objetos
for index, row in df.iterrows():
    matricula = row['matricula']
    data = row['data_hora_iso']

    # Cabeça e corpo do arquivo Json

    headers = {
        "identifier": "90.254.217/0001-10",
        "key": "67a4d6fa-ff21-487c-9426-7e22bb8a9142",
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
        "Matricula": matricula,
        "DataHoraApontamento": data,
        "ResponseType": "AS400V1"
    }
    print(data)
    print("Requisição enviada")
    print(payload)

    # Requisição Post no endpoint
    response = requests.post(
        "https://www.dimepkairos.com.br/RestServiceApi/Mark/SetMarks",
        json=payload,
        headers=headers)

    # Impressão da Resposta
    print('Resposta da API')
    print(response.text)

    if response.status_code == 200:
        print(f'Requisição para {matricula} enviada com sucesso')
    else:
        print(f'Erro ao enviar a requisição para {matricula} {response.text}')

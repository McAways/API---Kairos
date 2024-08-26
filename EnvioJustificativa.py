import pandas as pd
import requests

# Leitura do arquivo Excel
df = pd.read_excel(r'C:\Users\WI140\Desktop\TESTE JUST\Panda\API.xlsx')

for index, row in df.iterrows():
    idjust = row['Id-justificativa']
    idusu = row['Id-usuario']
    idfunc = row['Id-funcionario']
    # Converte Timestamp para string no formato YYYY-MM-DD
    data = row['data'].strftime('%Y-%m-%d')
    # Converte time para string no formato HH:MM
    horas = row['horas'].strftime('%H:%M')

    headers = {
        'identifier': "04.689.450/0001-94",
        'key': "319c0905-9f40-429e-b775-c243dea4e617",
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
        "IdJustification": idjust,
        "IdUser": idusu,
        "IdEmployee": idfunc,
        "QtdHours": horas,
        "Date": data,
        "Notes": "Enviado via API",
        "RequestType": "1",
        "ResponseType": "AS400V1"
    }

    print("Enviado Payload:")
    print(payload)

    response = requests.post(
        'https://www.dimepkairos.com.br/RestServiceApi/PreJustificationRequest/PreJustificationRequest',
        json=payload,
        headers=headers
    )

    print("Resposta da API:")
    print(response.text)

    if response.status_code == 200:
        print("Sucesso")
    else:
        print(f"Falha: {response.status_code}")

import os
import pandas as pd
import requests

# Leitura do arquivo Excel
df = pd.read_excel(r'C:\Users\WI140\Desktop\TESTE JUST\Panda\API.xlsx')

#Lista de armazenamento dos resultados
resultados = []

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

    # Analisando a resposta da API
    try:
        response_json = response.json()
        if response_json.get("Sucesso"):
            status = "Sucesso"
            mensagem = response_json.get("Mensagem", "")
        else:
            status = "Falha"
            mensagem = response_json.get("Mensagem", "Erro desconhecido")
    except ValueError:
        # Caso a resposta não seja um JSON válido
        status = "Falha"
        mensagem = "Resposta inválida da API"
        
    # Salvando os detalhes da resposta
    resultados.append([idfunc, data, status, mensagem])
    
# Criando um DataFrame com os resultados
df_resultados = pd.DataFrame(
    resultados, columns=["Matricula", "DataHoraApontamento", "Status", "Mensagem"])


# Salvando o resultado em um arquivo Excel
output_path = r'C:\Users\luis.marques\Desktop\VsCode\resultado.xlsx'

if os.path.exists(output_path):
    os.remove(output_path)

df_resultados.to_excel(output_path, index=False)

print(f'Resultados salvos em: {output_path}')

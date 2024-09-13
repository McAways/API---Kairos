import requests
import pandas as pd
import openpyxl
import os

# Leitor do arquivo excel (mudar o diretório dependendo do pc)
df = pd.read_excel(r'C:\Users\WI140\Desktop\TESTE JUST\Marcacaoteste.xlsx')
df['data_hora'] = pd.to_datetime(df['data'].astype(str) + ' ' +
                                 df['hora'].astype(str),
                                 format='mixed')

df['data_hora_iso'] = df['data_hora'].dt.strftime('%d/%m/%Y %H:%M')

# Lista para armazenar os resultados
resultados = []

# Transformando a leitura das colunas em objetos
for index, row in df.iterrows():
    matricula = row['matricula']
    data = row['data_hora_iso']

    # Cabeça e corpo do arquivo Json
    headers = {
        "identifier": "68.839.228/0001-03",
        "key": "67a4d6fa-ff21-487c-9426-7e22bb8a9142",
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
        "Matricula": matricula,
        "DataHoraApontamento": data,
        "ResponseType": "AS400V1"
    }

    # Requisição Post no endpoint
    response = requests.post(
        "https://www.dimepkairos.com.br/RestServiceApi/Mark/SetMarks",
        json=payload,
        headers=headers)

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
    resultados.append([matricula, data, status, mensagem])

# Criando um DataFrame com os resultados
df_resultados = pd.DataFrame(
    resultados, columns=["Matricula", "DataHoraApontamento", "Status", "Mensagem"])


# Salvando o resultado em um arquivo Excel
output_path = r'C:\Users\WI140\Desktop\TESTE JUST\resultados_envio.xlsx'

if os.path.exists(output_path):
    os.remove(output_path)

df_resultados.to_excel(output_path, index=False)

print(f'Resultados salvos em: {output_path}')

import requests
import pandas as pd
import openpyxl
import os
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import NamedStyle

# Leitor do arquivo excel (mudar o diretório dependendo do pc)
df = pd.read_excel(r'C:\Users\WI140\Documents\Arquivos Gerados\marc.xlsx')

# Corrigindo a leitura de datas, assumindo que o primeiro valor é o dia
df['data_hora'] = pd.to_datetime(df['data'].astype(str) + ' ' +
                                 df['hora'].astype(str),
                                 dayfirst=True, errors='coerce')

# Formatando para ISO com o formato correto (DD/MM/YYYY)
df['data_hora_iso'] = df['data_hora'].dt.strftime('%d/%m/%Y %H:%M')

# Lista para armazenar os resultados
resultados = []

# Transformando a leitura das colunas em objetos
for index, row in df.iterrows():
    matricula = row['matri']
    data = row['data_hora_iso']

    # Cabeça e corpo do arquivo Json
    headers = {
        "identifier": "53.484.549/0001-65",
        "key": "48c7b084-f51d-4225-8067-98923b365aed",
        "cpf": "69306288085",
        'User-Agent': 'PostmanRuntime/7.30.0'
    }

    payload = {
        "Matricula": matricula,
        "DataHoraApontamento": data,
        "CpfResponsavel": "69306288085",
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
output_path = r'C:\Users\WI140\Documents\Arquivos Gerados\resultado.xlsx'

if os.path.exists(output_path):
    os.remove(output_path)

# Exportando o DataFrame para Excel sem o argumento date_format
df_resultados.to_excel(output_path, index=False)

# Ajustando o formato da coluna DataHoraApontamento com openpyxl
wb = openpyxl.load_workbook(output_path)
ws = wb.active

# Criando um estilo de data
date_style = NamedStyle(name="datetime", number_format="DD/MM/YYYY HH:MM")

# Aplicando o estilo na coluna DataHoraApontamento
for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
    for cell in row:
        cell.style = date_style

# Salvando o arquivo Excel com o formato corrigido
wb.save(output_path)

print(f'Resultados salvos em: {output_path}')

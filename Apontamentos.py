import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


def apply_borders(ws):
    # Define o estilo da borda
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Aplica a borda a todas as células do worksheet
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

# Função para fazer a requisição e salvar os dados filtrados no Excel
def get_filtered_data_and_save_to_excel(url, payload, headers, output_path):
    response = requests.post(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            # Decodifica o JSON
            data = response.json()

            # Verifica se o campo "Obj" existe e contém dados
            if "Obj" in data and isinstance(data['Obj'], list):
                # Lista para armazenar os dados das entradas
                all_entries = []
                # Lista para armazenar dados fixos
                all_fixed_data = []

                # Itera sobre cada objeto no campo "Obj"
                for item in data['Obj']:
                
                    # Adiciona os dados fixos (não relacionados às entradas) em uma lista
                    fixed_data = {
                        'Empresa': item['InfoEmpresa']['Nome'],
                        'CNPJ': item['InfoEmpresa']['CNPJCPF'],
                        'PIS': item['InfoFuncionario']['PIS'],
                        'Funcionario': item['InfoFuncionario']['Nome'],
                        'Matricula': item['InfoFuncionario']['Matricula']
                    }
                    # Adiciona os dados fixos à lista de dados
                    all_fixed_data.append(fixed_data)
                    

                    # Itera sobre cada entrada dentro do campo 'Entradas'
                    for entrada in item['Entradas']:
                        # Cria um dicionário para cada entrada com as informações necessárias
                        entry_data = {
                            'Data': entrada['Data'],
                            'Horario': entrada['Horario'],
                            'Apontamentos': entrada['Apontamentos'],
                            'Horas Trabalhadas': entrada['HTrab'],
                            'Horas Extras': entrada['HE'],
                            'Descontos': entrada['Descontos'],
                            'Debito': entrada['Debito'],
                            'Credito': entrada['Credito']
                        }
                        # Adiciona o dicionário à lista de todas as entradas
                        all_entries.append(entry_data)
                        
                

                # Converte as listas de entradas e dados fixos em DataFrames
                df_entries = pd.DataFrame(all_entries)
                df_fixed = pd.DataFrame(all_fixed_data)

                # Combine os dois DataFrames
                # Para garantir que o número de linhas em df_fixed seja igual ao de df_entries, replicamos os valores fixos
                df_fixed_repeated = df_fixed.loc[df_fixed.index.repeat(len(df_entries)//len(df_fixed))].reset_index(drop=True)

                # Combina os DataFrames lado a lado
                final_df = pd.concat([df_fixed_repeated, df_entries], axis=1)
                
                # Adiciona a nova coluna "Ajustes" com valores em branco ou padrão
                final_df['Justificativas'] = ''  # ou você pode definir um valor padrão como: 'Ajuste Pendente'
                final_df['Entrada'] = ''
                final_df['Saida Pausa'] = ''
                final_df['Volta Pausa'] = ''
                final_df['Saida'] = ''

                # Cria um Workbook do openpyxl
                wb = Workbook()
                ws = wb.active
                
                justificativa_options = ['Atestado', 'Folga', 'Justificativa de Horas', 'Abono']
                justificativa_str = ','.join(justificativa_options)
                
                justificativa_validation = DataValidation(type='list', formula1=f'"{justificativa_str}"', allow_blank=True)
                justificativa_validation.error = 'Escolher valores da lista'
                justificativa_validation.errorTitle = 'Entrada Invalida'
                justificativa_validation.prompt = 'Selecione uma justificativa'
                justificativa_validation.promptTitle = 'Justificativas'
                
                justificativa_col_index = final_df.columns.get_loc("Justificativas") + 1
                justificativa_col_letter = ws.cell(row=1, column=justificativa_col_index).column_letter
                
                ws.add_data_validation(justificativa_validation)
                justificativa_validation.add(f'{justificativa_col_letter}2:{justificativa_col_letter}{len(final_df)+1}')

                # Escreve o DataFrame no Excel usando openpyxl
                for r_idx, row in enumerate(dataframe_to_rows(final_df, index=False, header=True), 1):
                    for c_idx, value in enumerate(row, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value)

                # Ajusta a largura das colunas automaticamente
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column].width = adjusted_width

                # Formatação de cor alternada nas linhas
                fill_grey = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
                fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                
                

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    for cell in row:
                        # Alterna entre as cores
                        if cell.row % 2 == 0:
                            cell.fill = fill_grey
                        else:
                            cell.fill = fill_white
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        
                apply_borders(ws)

                # Salva o arquivo Excel com formatação
                wb.save(output_path)
                print(f"Arquivo Excel salvo com sucesso em {output_path}")
            else:
                print("Nenhum dado no campo 'Obj'.")
        except ValueError as e:
            print(f"Erro ao decodificar JSON: {e}")
            print("Conteúdo da resposta:")
            print(response.text)
    else:
        print(f"Falha: {response.status_code}")
        print(response.text)
datafim = '24/09/2024'

# URL do endpoint e payload
url = 'https://www.dimepkairos.com.br/RestServiceApi/ReportEmployeePunch/GetReportEmployeePunch'  # Modifique para o endpoint correto
payload = {"param": "value"}  # Modifique conforme necessário

# Headers da requisição
headers = {
        "identifier": "29.024.277/0001-36",
        "key": "a06a8bc3-9b7f-4d97-b45d-15549eee8063",
        'User-Agent': 'PostmanRuntime/7.30.0'
}

payload = {
        "MatriculaPessoa": [],
        "DataInicio":"16/09/2024",
        "DataFim":"24/09/2024",
        "ResponseType":"AS400V1"
}

# Caminho de saída do arquivo Excel
output_path = r'C:\Users\luis.marques\Desktop\VsCode\Tesdte.xlsx'

# Chama a função para obter e filtrar os dados e salvar no Excel
get_filtered_data_and_save_to_excel(url, payload, headers, output_path)

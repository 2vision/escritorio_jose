import gspread
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os

def executar():
    base_pf = acessar_sheets('PF')
    base_pj = acessar_sheets('PJ')

    processos_pf = capturar_processos_pf(base_pf)
    processos_pj = capturar_processos_pj(base_pj)

    for processo in processos_pf:
        gerar_documento(processo, "PF")

    for processo in processos_pj:
        gerar_documento(processo, "PJ")


def acessar_sheets(aba):
    url_a_partir_do_d = '1hwUdFxkPQkdovOGpmMt8cqXbMnZoMFsFUKGefL8-bOU'
    google_cloud = gspread.service_account(filename='mtadv-449314-47f9a9de429d.json')
    sheet = google_cloud.open_by_key(url_a_partir_do_d)
    return sheet.worksheet(aba)


def capturar_processos_pf(base):
    df = pd.DataFrame(base.get_all_values())
    df.columns = df.iloc[0]  # Define a primeira linha como cabeçalho
    df = df[1:]  # Remove a linha do cabeçalho dos dados
    df.reset_index(drop=True, inplace=True)

    processos = []
    for _, row in df.iterrows():
        processo = {
            'nome': row['Nome'],
            'nacionalidade': row['Nacionalidade'],
            'cpf': row['CPF'],
            'endereco': row['Endereço'],
            'numero': row['Número'],
            'complemento': row['Complemento'],
            'bairro': row['Bairro'],
            'cidade': row['Cidade'],
            'estado': row['Estado'],
            'cep': row['CEP'],
            'email': row['e-mail'],
            'processo': row['Processo'],
            'numero_vara': row['nº Vara'],
            'competencia': row['Competência'],
            'ente': row['Ente'],
            'jurisdicao': row['Jurisdição'],
            'valor_causa': row['Valor da causa'],
            'valor_honorarios': row['Valor de honorários iniciais Númeral'],
            'parcelas': row['Parcelas'],
            'data_documento': row['Data do documento']
        }
        processos.append(processo)
    return processos


def capturar_processos_pj(base):
    df = pd.DataFrame(base.get_all_values())
    df.columns = df.iloc[0]  # Define a primeira linha como cabeçalho
    df = df[1:]  # Remove a linha do cabeçalho dos dados
    df.reset_index(drop=True, inplace=True)

    processos = []
    for _, row in df.iterrows():
        processo = {
            'nome_empresa': row['Nome da empresa'],
            'cnpj': row['CNPJ'],
            'endereco_empresa': row['Endereço'],
            'numero_empresa': row['Número'],
            'complemento_empresa': row['Complemento'],
            'bairro_empresa': row['Bairro'],
            'cidade_empresa': row['Cidade'],
            'estado_empresa': row['Estado'],
            'cep_empresa': row['CEP'],
            'nome_representante': row['Nome do representante'],
            'nacionalidade_representante': row['Nacionalidade'],
            'cpf_representante': row['CPF'],
            'endereco_representante': row['Endereço'],
            'numero_representante': row['Número'],
            'complemento_representante': row['Complemento'],
            'bairro_representante': row['Bairro'],
            'cidade_representante': row['Cidade'],
            'estado_representante': row['Estado'],
            'cep_representante': row['CEP'],
            'email': row['e-mail'],
            'processo': row['Processo'],
            'numero_vara': row['nº Vara'],
            'competencia': row['Competência'],
            'ente': row['Ente'],
            'valor_causa': row['Valor da causa'],
            'valor_honorarios': row['Valor de honorários iniciais'],
            'numerar': row['Númeral'],
            'valor_honorarios_escrito': row['Valor honorários escrito'],
            'parcelas': row['Parcelas'],
            'valor_parcela_numeral': row['Valor da parcela Numeral'],
            'valor_parcela_extenso': row['Valor parcela por extenso'],
            'data_documento': row['Data do documento']
        }
        processos.append(processo)
    return processos


def arquivos_word():

    base_1 = '10-QYHcIK6ORUnGsSKx8apLXrh19_M992'
    base_2 = '1Dgh8aDjulRXiXvwrUSmdL2cfxSXJ0D0P'
    procuracao_pf = '1xUpqZgnigFyc-NQb49hqyhkZOjbiZ2XL'
    procuracao_pj = '1pjJikJa1zYmuvORHCiC1jZ68h3L17o0U'

def gerar_documento(processo, tipo_processo):
    SCOPES = ["https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/drive"]
    SERVICE_ACCOUNT_FILE = "mtadv-449314-71b3080b64d1.json"

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    service = build("docs", "v1", credentials=credentials)
    drive_service = build("drive", "v3", credentials=credentials)

    DOCUMENT_MODELO_ID = "1QviBHMz-4kWPSvSl1bSbbzH90IS6guStmP27QMxAntM"

    # Criar uma pasta com o nome do cliente
    nome_pasta = processo['nome'] if tipo_processo == 'PF' else processo['nome_empresa']
    pasta_metadata = {
        'name': nome_pasta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    pasta = drive_service.files().create(body=pasta_metadata, fields='id').execute()
    pasta_id = pasta.get('id')

    # Mover a pasta para o seu Drive (sem remover da raiz da conta de serviço)
    pasta_principal_id = '1BJ1HjguP5Z8QaLkTp85-oJtJqTiwMdJi'  # Substitua pelo ID da pasta principal no seu Drive
    drive_service.files().update(
        fileId=pasta_id,
        addParents=pasta_principal_id,  # Coloca a pasta dentro de uma pasta específica no seu Drive
        fields='id, parents'
    ).execute()

    # Criar o documento
    novo_documento = drive_service.files().copy(
        fileId=DOCUMENT_MODELO_ID,
        body={'name': f"Contrato_{tipo_processo}_{nome_pasta}"}
    ).execute()

    novo_documento_id = novo_documento['id']

    # Mover o documento para a pasta criada
    drive_service.files().update(
        fileId=novo_documento_id,
        addParents=pasta_id,  # Adiciona o documento na pasta criada
        fields='id, parents'
    ).execute()

    # Compartilhar o documento
    compartilhar_doc(drive_service, novo_documento_id)

    requests = []
    for key, value in processo.items():
        requests.append({
            'replaceAllText': {
                'containsText': {'text': f'{{{key}}}', 'matchCase': 'true'},
                'replaceText': value
            }
        })

    service.documents().batchUpdate(
        documentId=novo_documento_id,
        body={'requests': requests}
    ).execute()

    print(f"Documento gerado com sucesso: {novo_documento['name']} na pasta {nome_pasta}")

def compartilhar_doc(drive_service, documento_id):

    permissions = drive_service.permissions()

    permissions.create(
        fileId=documento_id,
        body={
            'type': 'anyone',
            'role': 'writer'
        },
        fields='id'
    ).execute()

executar()
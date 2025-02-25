import gspread
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from num2words import num2words

# IDs dos documentos no Google Drive
DOCUMENTOS = {
    'CONTRATO_PF': '1QviBHMz-4kWPSvSl1bSbbzH90IS6guStmP27QMxAntM',
    'CONTRATO_PJ': '1NKwZ-xM0ihiYj5Ry8xzD7DyX8_ZdJFHIKkHTCPwOpTQ',
    'PROCURACAO_PF': '1Lz_UZFbYzhpdAz1TxdVMlA1V8LKE6LhHhwCetUBiibU',
    'PROCURACAO_PJ': '1r7niFxEbD9RdXCoahMYnQimZVRMpB0qbKRpfplgSZWM',
}
PASTA_GERAL_ID = '1BJ1HjguP5Z8QaLkTp85-oJtJqTiwMdJi'
SHEET_ID = '1hwUdFxkPQkdovOGpmMt8cqXbMnZoMFsFUKGefL8-bOU'

# Configuração de autenticação
CREDENTIALS_FILE = 'mtadv-449314-47f9a9de429d.json'
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']


def executar():
    planilhas = {
        'PF': carregar_dados_sheets('PF'),
        'PJ': carregar_dados_sheets('PJ'),
    }

    for tipo, dados in planilhas.items():
        processos = processar_dados(dados, tipo)
        for processo in processos:
            gerar_documentos(processo, tipo)


def carregar_dados_sheets(aba):
    cliente = gspread.service_account(filename=CREDENTIALS_FILE)
    planilha = cliente.open_by_key(SHEET_ID)
    return planilha.worksheet(aba)


def processar_dados(sheet, tipo):
    df = pd.DataFrame(sheet.get_all_values())
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    processos = []

    for _, row in df.iterrows():
        valor_honorarios = converter_float(row['Valor de honorários iniciais Númeral'])
        parcelas = converter_int(row['Parcelas'])
        valor_parcela = valor_honorarios / parcelas

        if tipo == 'PF':
            processo = {
                'nome': str(row['Nome']).upper(),
                'nacionalidade': row['Nacionalidade'],
                'cpf': formatar_cpf(row['CPF']),
                'endereco': row['Endereço'].title(),
                'numero_endereco': row['Número'],
                'complemento': row['Complemento'],
                'bairro': row['Bairro'].title(),
                'cidade': row['Cidade'].title(),
                'estado': row['Estado'],
                'cep': formatar_cep(row['CEP']),
                'email': row['e-mail'],
                'processo': row['Processo'],
                'numero_vara': row['nº Vara'],
                'competencia': row['Competência'],
                'ente': str(row['Ente']).upper(),
                'jurisdicao': row['Jurisdição'].title(),
                'valor_causa': formatar_valor(row['Valor da causa']),
                'valor_honorarios_iniciais': formatar_valor(row['Valor de honorários iniciais Númeral']),
                'valor_honorarios_iniciais_extenso': valor_por_extenso(
                    converter_float(row['Valor de honorários iniciais Númeral'])),
                'parcelas_honorarios_iniciais': row['Parcelas'],
                'parcelas_honorarios_iniciais_extenso': valor_por_extenso(row['Parcelas'], 'int'),
                'data_do_documento': row['Data do documento'],
                'valor_parcelas': formatar_valor(valor_parcela),
                'valor_parcelas_extenso': valor_por_extenso(valor_parcela),
            }
        else:
            processo = {
                'nome_da_empresa': str(row['Nome da empresa']).upper(),
                'cnpj': formatar_cnpj(row['CNPJ']),
                'endereco_pj': row['Endereço'].title(),
                'numero_endereco_pj': row['Número'],
                'complemento_pj': row['Complemento'],
                'bairro_pj': row['Bairro'].title(),
                'cidade_pj': row['Cidade'].title(),
                'estado_pj': row['Estado'],
                'cep_pj': formatar_cep(row['CEP']),
                'nome_representante': str(row['Nome do representante']).upper(),
                'nacionalidade_representante': row['Nacionalidade'],
                'cpf_representante': formatar_cpf(row['CPF']),
                'endereco_representante': row['Endereço Representante'].title(),
                'numero_endereco_representante': row['Número Representante'],
                'complemento_representante': row['Complemento Representante'],
                'bairro_representante': row['Bairro Representante'].title(),
                'cidade_representante': row['Cidade Representante'].title(),
                'estado_representante': row['Estado Representante'],
                'cep_representante': formatar_cep(row['CEP Representante']),
                'email': row['e-mail'],
                'processo': row['Processo'],
                'vara': row['nº Vara'],
                'jurisdicao': row['Jurisdição'].title(),
                'competencia': row['Competência'],
                'ente': str(row['Ente']).upper(),
                'valor_causa': formatar_valor(row['Valor da causa']),
                'valor_honorarios': formatar_valor(row['Valor de honorários iniciais Númeral']),
                'valor_honorarios_extenso': valor_por_extenso(
                    converter_float(row['Valor de honorários iniciais Númeral'])),
                'parcelas': row['Parcelas'],
                'parcelas_extenso': valor_por_extenso(row['Parcelas'], 'int'),
                'valor_parcelas': formatar_valor(valor_parcela),
                'valor_parcelas_extenso': valor_por_extenso(valor_parcela),
                'data_do_documento': row['Data do documento']
            }

        processos.append(processo)

    return processos


def gerar_documentos(processo, tipo):
    credenciais = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    drive_service = build('drive', 'v3', credentials=credenciais)
    docs_service = build('docs', 'v1', credentials=credenciais)

    nome_pasta = processo.get('nome') or processo.get('nome_da_empresa')
    pasta_id = criar_pasta(drive_service, nome_pasta)

    contrato_id, procuracao_id = DOCUMENTOS[f'CONTRATO_{tipo}'], DOCUMENTOS[f'PROCURACAO_{tipo}']
    gerar_doc_drive(drive_service, docs_service, contrato_id, processo, pasta_id, f'Contrato_{tipo}_{nome_pasta}')
    gerar_doc_drive(drive_service, docs_service, procuracao_id, processo, pasta_id, f'Procuracao_{tipo}_{nome_pasta}')


def criar_pasta(drive_service, nome_pasta):
    metadata = {'name': nome_pasta, 'mimeType': 'application/vnd.google-apps.folder'}
    pasta = drive_service.files().create(body=metadata, fields='id').execute()
    drive_service.files().update(fileId=pasta['id'], addParents=PASTA_GERAL_ID).execute()
    return pasta['id']


def gerar_doc_drive(drive_service, docs_service, modelo_id, dados, pasta_id, nome_arquivo):
    novo_doc = drive_service.files().copy(fileId=modelo_id, body={'name': nome_arquivo}).execute()
    novo_doc_id = novo_doc['id']
    drive_service.files().update(fileId=novo_doc_id, addParents=pasta_id).execute()

    drive_service.permissions().create(
        fileId=novo_doc_id,
        body={'type': 'anyone', 'role': 'writer'},
        fields='id'
    ).execute()

    requests = [{
        'replaceAllText': {
            'containsText': {'text': f'{{{key}}}', 'matchCase': True},
            'replaceText': str(value)
        }
    } for key, value in dados.items()]

    docs_service.documents().batchUpdate(documentId=novo_doc_id, body={'requests': requests}).execute()
    print(f'Documento gerado: {nome_arquivo}')


def converter_float(valor):
    if valor:
        valor = valor.replace('.', '').replace(',', '.')
        return float(valor)
    return 0


def converter_int(valor):
    if valor:
        return int(valor)
    return 0


def formatar_valor(valor):
    if valor:
        if isinstance(valor, str):
            valor = float(valor.replace('.', '').replace(',', '.'))
            return f'{valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
        if isinstance(valor, float):
            return f'{valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return 0


def valor_por_extenso(valor, tipo='float'):
    if valor:
        valor_texto = num2words(valor, lang='pt_BR', to='currency')
        valor_texto = valor_texto.replace('euro', 'real').replace('euros', 'reais')

        if tipo == 'int':
            valor_texto = valor_texto.replace('reais', '').replace('real', '').strip()

        return valor_texto


def formatar_cpf(valor):
    if valor:
        valor = str(valor).replace('.', '').replace('-', '')
        cpf = "{:011}".format(int(valor))
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    return 0


def formatar_cnpj(valor):
    if valor:
        valor = str(valor).replace('.', '').replace('-', '').replace('/', '')
        cnpj = "{:014}".format(int(valor))
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return 0


def formatar_cep(valor):
    if valor:
        valor = str(valor).replace('.', '').replace('-', '')
        cep = "{:08}".format(int(valor))
        return f"{cep[:5]}-{cep[5:]}"
    return 0


def formatar_processo(valor):
    if valor:
        valor = str(valor).replace('.', '').replace('-', '')
        processo = "{:020}".format(int(valor))
        return f"{processo[:7]}-{processo[7:9]}.{processo[9:13]}.{processo[13:14]}.{processo[14:16]}.{processo[16:20]}"
    return 0


executar()

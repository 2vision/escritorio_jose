import locale
from datetime import datetime

import gspread
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from num2words import num2words

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

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

    registros_executados = {'PF': [], 'PJ': []}

    for tipo, dados in planilhas.items():
        processos, registros_executados = processar_dados(dados, tipo, registros_executados)
        for processo in processos:
            gerar_documentos(processo, tipo)

    atualizar_planilha(planilhas, registros_executados)


def carregar_dados_sheets(aba):
    cliente = gspread.service_account(filename=CREDENTIALS_FILE)
    planilha = cliente.open_by_key(SHEET_ID)
    return planilha.worksheet(aba)


def processar_dados(sheet, tipo, registros_executados):
    df = pd.DataFrame(sheet.get_all_values())
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    processos = []

    for index, row in df.iterrows():
        if row.get('Gerado', '').strip().lower() == 'sim':
            continue

        valor_honorarios = converter_float(row['Valor de honorários iniciais Númeral'])
        valor_entrada = converter_float(row['Valor da entrada'])
        parcelas = converter_int(row['Parcelas'])

        if valor_entrada != 0:
            valor_parcela = (valor_honorarios - valor_entrada) / parcelas
        else:
            valor_parcela = valor_honorarios / parcelas

        dados_comuns = {
            'endereco': row['Endereço'].title(),
            'numero_endereco': row['Número'],
            'complemento': row['Complemento'],
            'bairro': row['Bairro'].title(),
            'cidade': row['Cidade'].title(),
            'estado': row['Estado'].upper(),
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
                converter_float(row['Valor de honorários iniciais Númeral'])
            ),
            'valor_entrada': formatar_valor(row['Valor da entrada']),
            'valor_entrada_extenso': valor_por_extenso(row['Valor da entrada']),
            'parcelas_honorarios_iniciais': row['Parcelas'],
            'parcelas_honorarios_iniciais_extenso': valor_por_extenso(row['Parcelas'], 'int'),
            'valor_parcelas': formatar_valor(valor_parcela),
            'valor_parcelas_extenso': valor_por_extenso(valor_parcela),
            'data_do_documento': datetime.today().strftime('%d de %B de %Y'),
            'dia_do_boleto': row["Boleto dia vencimento"]
        }

        if tipo == 'PF':
            processo = {
                **dados_comuns,
                'nome': str(row['Nome']).upper(),
                'nacionalidade': row['Nacionalidade'],
                'cpf': formatar_cpf(row['CPF']),
            }
        else:
            processo = {
                **dados_comuns,
                'vara': row['nº Vara'],
                'nome_da_empresa': str(row['Nome da empresa']).upper(),
                'cnpj': formatar_cnpj(row['CNPJ']),
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
            }

        processos.append(processo)
        registros_executados[tipo].append(index + 2)

    return processos, registros_executados


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

    document = docs_service.documents().get(documentId=novo_doc_id).execute()

    for elemento in document.get('body').get('content'):
        if 'paragraph' in elemento:
            paragrafos = elemento['paragraph']

    if 'Contrato' in nome_arquivo:
        tem_boleto = dados.get('dia_do_boleto')
        if 'PF' in nome_arquivo:
            start_index, end_index = (4523, 5173) if tem_boleto else (3858, 4523)
        elif 'PJ' in nome_arquivo:
            start_index, end_index = (4881, 5531) if tem_boleto else (4216, 4881)

        requests = [{
            'deleteContentRange': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                }
            }
        }]
        docs_service.documents().batchUpdate(documentId=novo_doc_id, body={'requests': requests}).execute()

    requests = [{
        'replaceAllText': {
            'containsText': {'text': f'{{{key}}}', 'matchCase': True},
            'replaceText': str(value)
        }
    } for key, value in dados.items()]

    docs_service.documents().batchUpdate(documentId=novo_doc_id, body={'requests': requests}).execute()
    print(f'Documento gerado: {nome_arquivo}')


def atualizar_planilha(planilhas, registros_executados):
    for tipo, sheet in planilhas.items():
        updates = []
        for linha in registros_executados[tipo]:
            if tipo == 'PF':
                coluna = 'V'
                updates.append({
                    "range": f"{coluna}{linha}",
                    "values": [["Sim"]]
                })
            elif tipo == 'PJ':
                coluna = 'AE'
                updates.append({
                    "range": f"{coluna}{linha}",
                    "values": [["Sim"]]
                })

        if updates:
            sheet.batch_update(updates)


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

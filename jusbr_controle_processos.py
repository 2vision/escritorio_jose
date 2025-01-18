import json
import locale
import os
from datetime import datetime
import gspread
import pandas as pd
import requests
from datetime import datetime, timedelta

NOME_ARQUIVO_PARA_SALVAR = 'andamentos_processos'
BEARER_CODE = 'eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3MzcyNTUzNTksImlhdCI6MTczNzIyNjU2MCwiYXV0aF90aW1lIjoxNzM3MjI2Mjg3LCJqdGkiOiIwM2UzYzljNy1mMjhkLTQ5MzQtODFjMS1kN2RhOGNjN2VkYzkiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6ImFjY291bnQiLCJzdWIiOiI4ZDBjM2JjYy0zZDliLTRmZTMtOGU4Yy1hYTdjNDM5OTRhMGIiLCJ0eXAiOiJCZWFyZXIiLCJhenAiOiJwb3J0YWxleHRlcm5vLWZyb250ZW5kIiwibm9uY2UiOiJhMzM2OTgzZC1mZTdlLTRlOGYtOTZjMS0xOTc2MTY0ZGRmMDUiLCJzZXNzaW9uX3N0YXRlIjoiZjRlYjQyMmEtZDIxNS00MjdjLTk2OGUtOTA2YWMzODFhMWJhIiwiYWNyIjoiMCIsImFsbG93ZWQtb3JpZ2lucyI6WyJodHRwczovL3BvcnRhbGRlc2Vydmljb3MucGRwai5qdXMuYnIiXSwicmVzb3VyY2VfYWNjZXNzIjp7ImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJmNGViNDIyYS1kMjE1LTQyN2MtOTY4ZS05MDZhYzM4MWExYmEiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJFRFVBUkRPIFBFUkVJUkEgR09NRVMiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiI3Njc4NzA0NDAyMCIsImdpdmVuX25hbWUiOiJFRFVBUkRPIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkdPTUVTIiwiY29ycG9yYXRpdm8iOlt7InNlcV91c3VhcmlvIjo1MzQ3MjA5LCJub21fdXN1YXJpbyI6IkVEVUFSRE8gUEVSRUlSQSBHT01FUyIsIm51bV9jcGYiOiI3Njc4NzA0NDAyMCIsInNpZ190aXBvX2NhcmdvIjoiQURWIiwiZmxnX2F0aXZvIjoiUyIsInNlcV9zaXN0ZW1hIjowLCJzZXFfcGVyZmlsIjowLCJkc2Nfb3JnYW8iOiJPQUIiLCJzZXFfdHJpYnVuYWxfcGFpIjowLCJkc2NfZW1haWwiOiJzZWNyZXRhcmlhQGVkdWFyZG9nb21lcy5hZHYuYnIiLCJzZXFfb3JnYW9fZXh0ZXJubyI6MCwiZHNjX29yZ2FvX2V4dGVybm8iOiJPQUIiLCJvYWIiOiJSUzkxNjMxIn1dLCJlbWFpbCI6InNlY3JldGFyaWFAZWR1YXJkb2dvbWVzLmFkdi5iciJ9.LKMs90hzI3gHZXwNBugxA09Uwt_YWvjihcv4CaNqPYqOnhdZlId2g3sl9dWjvyNxLgoLNEHfWCrb_u4UVCoRirMd1Go1zA36ibyppd9OqzQHvB4Ec0E8ZAfvcEzstkxhAkB9RqYDOX5dn7TPTIaw6FAbiRmh_51ZRKsS4qEQ5h5PqwO9DnMS75GqaLR1g9DqdvgopV9v2EZp8e57F3OYg0mfGMTYW8h4yBNs6WijfL-yNdLLBLZZlDEh2j5CCfKRTI4CyKTz6A87yUV5SJKOFh5ubv7_mUsIO7tY4w3hdb0cMT4fP3-207hsODkVjWMbowQ-WPMYohAOH_9kFk8E0w'


def executar():
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    processos = capturar_processos_sheets()

    for processo in processos:
        dados = api_jusbr(processo.get('numero_processo'))
        dados_prontos, movimentos = dados_formatados(dados, processo.get('numero_processo'), processo.get('ultimo_movimento'))
        salvar_informacoes_no_json(dados_prontos)
        salvar_informacoes_no_json(movimentos)
    salvar_informacoes_no_excel()


def capturar_processos_sheets():
    processos_consulta_andamento = acessar_sheets('Procesos_Andamentos')
    df = pd.DataFrame(processos_consulta_andamento.get_all_records())
    processos = []

    for index, linha in df.iterrows():
        numero_processo = linha['Número do processo']
        ultimo_movimento = linha['Data último movimento']
        processos.append({'numero_processo': numero_processo,
                          'ultimo_movimento': ultimo_movimento
                          })
    return processos


def api_jusbr(processo):
    processo_limpo = processo.replace('.', '').replace('-', '')
    url_base = f'https://portaldeservicos.pdpj.jus.br/api/v2/processos/{processo_limpo}'
    headers = {'Authorization': f'Bearer {BEARER_CODE}'}
    response = requests.get(url_base, headers=headers)

    if response and response.status_code == 200:
        return response.json()


def dados_formatados(dados, numero_processo, ultimo_movimento):
    proximo_movimento_atual = None
    processo = dados[0].get('tramitacaoAtual')
    valor_acao = processo.get('valorAcao')
    grau = processo.get('grau').get('nome')
    tribunal = processo.get('tribunal').get('sigla')
    classe_principal = processo.get('classe')[0].get('descricao')
    data_de_distribuicao = datetime.strptime(processo.get('distribuicao')[0].get('dataHora'), '%Y-%m-%dT%H:%M:%S.%f')

    ultima_semana = datetime.now() - timedelta(days=7)
    movimentos = []

    for movimento in processo.get('movimentos'):

        movimento_data = datetime.strptime(movimento.get('dataHora'), '%Y-%m-%dT%H:%M:%S.%f')

        ultimo_movimento = datetime.strptime(ultimo_movimento, '%Y-%m-%dT%H:%M:%S.%f')
        if movimento_data > ultima_semana and movimento_data > ultimo_movimento:
            descricao_movimento = movimento.get('descricao')

            if not proximo_movimento_atual and proximo_movimento_atual < movimento_data:
                proximo_movimento_atual = movimento_data

            movimento_data = movimento_data.strftime('%d/%m/%Y')
            movimentos.append([movimento_data, descricao_movimento])


    nome_ativo = None
    nome_passivo = None
    for parte in processo.get('partes'):

        if parte.get('polo') == 'ATIVO' and not nome_ativo:
            nome_ativo = parte.get('nome')

        if parte.get('polo') == 'PASSIVO' and not nome_passivo:
            nome_passivo = parte.get('nome')

    informacoes = []
    for movimento in movimentos:
        informacoes.append({
            'Processo': numero_processo,
            'Valor ação': locale.currency(valor_acao, grouping=True),
            'Estado': grau + ' ' + tribunal,
            'Classe Judicial': classe_principal,
            'Data de distribuição': data_de_distribuicao.strftime('%d/%m/%Y'),
            'Data movimento': movimento[0],
            'Movimento': movimento[1],
            'Polo ativo': nome_ativo.title(),
            'Polo passivo': nome_passivo.title()
        })

    informacoes_movimentos = {
                                'processo' : numero_processo,
                                'movimento' : proximo_movimento_atual
    }
    return informacoes, informacoes_movimentos


def salvar_informacoes_no_json(informacoes):
    if os.path.exists(f'{NOME_ARQUIVO_PARA_SALVAR}.json'):
        with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
            dados = json.load(f)
    else:
        dados = []

    dados.append(informacoes)

    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


def salvar_informacoes_no_excel():
    aba = acessar_sheets('TESTE')
    with open(f"{NOME_ARQUIVO_PARA_SALVAR}.json", "r", encoding='utf-8') as f:
        dados = json.load(f)

    dados_flat = [registro for grupo in dados for registro in grupo]
    df = pd.DataFrame(dados_flat)

    ultima_linha_preenchida = len(aba.get_all_values())
    linha_inicial = ultima_linha_preenchida + 1

    range_inicial = f"A{linha_inicial}"
    aba.update(range_inicial, df.values.tolist())

    if os.path.exists(f'{NOME_ARQUIVO_PARA_SALVAR}.json'):
        os.remove(f'{NOME_ARQUIVO_PARA_SALVAR}.json')

    print(f'o arquivo {NOME_ARQUIVO_PARA_SALVAR}.xlsx foi gerado com sucesso!')


def acessar_sheets(aba):
    url_apatir_do_d = '1AzdNQyAm_smi-t3tNwrMQHUiir42flWsMnJpdgVA6-k'
    google_cloud = gspread.service_account(filename='bancarioadvbox-e9fd143ea26d.json')
    sheet = google_cloud.open_by_key(url_apatir_do_d)
    return sheet.worksheet(aba)


executar()

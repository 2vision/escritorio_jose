import json
import locale
import os
import threading
from datetime import datetime, timedelta
from itertools import cycle

import gspread
import pandas as pd
import requests

NOME_ARQUIVO_PARA_SALVAR = 'andamentos_processos'
BEARER_CODES = cycle([
    "eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3MzcyNjk0ODcsImlhdCI6MTczNzI0MzMyMCwiYXV0aF90aW1lIjoxNzM3MjI2Mjg3LCJqdGkiOiJhNGU4MWM5NC1hMzdmLTRkOTItOGFhMS00MjYzMWU4ODU5NjIiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6ImFjY291bnQiLCJzdWIiOiI4ZDBjM2JjYy0zZDliLTRmZTMtOGU4Yy1hYTdjNDM5OTRhMGIiLCJ0eXAiOiJCZWFyZXIiLCJhenAiOiJwb3J0YWxleHRlcm5vLWZyb250ZW5kIiwibm9uY2UiOiJmNWFjYjRlOS1lYWRjLTQxYjktODg0NS0xM2I5NTlmYmM5ZjciLCJzZXNzaW9uX3N0YXRlIjoiZjRlYjQyMmEtZDIxNS00MjdjLTk2OGUtOTA2YWMzODFhMWJhIiwiYWNyIjoiMCIsImFsbG93ZWQtb3JpZ2lucyI6WyJodHRwczovL3BvcnRhbGRlc2Vydmljb3MucGRwai5qdXMuYnIiXSwicmVzb3VyY2VfYWNjZXNzIjp7ImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJmNGViNDIyYS1kMjE1LTQyN2MtOTY4ZS05MDZhYzM4MWExYmEiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJFRFVBUkRPIFBFUkVJUkEgR09NRVMiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiI3Njc4NzA0NDAyMCIsImdpdmVuX25hbWUiOiJFRFVBUkRPIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkdPTUVTIiwiY29ycG9yYXRpdm8iOlt7InNlcV91c3VhcmlvIjo1MzQ3MjA5LCJub21fdXN1YXJpbyI6IkVEVUFSRE8gUEVSRUlSQSBHT01FUyIsIm51bV9jcGYiOiI3Njc4NzA0NDAyMCIsInNpZ190aXBvX2NhcmdvIjoiQURWIiwiZmxnX2F0aXZvIjoiUyIsInNlcV9zaXN0ZW1hIjowLCJzZXFfcGVyZmlsIjowLCJkc2Nfb3JnYW8iOiJPQUIiLCJzZXFfdHJpYnVuYWxfcGFpIjowLCJkc2NfZW1haWwiOiJzZWNyZXRhcmlhQGVkdWFyZG9nb21lcy5hZHYuYnIiLCJzZXFfb3JnYW9fZXh0ZXJubyI6MCwiZHNjX29yZ2FvX2V4dGVybm8iOiJPQUIiLCJvYWIiOiJSUzkxNjMxIn1dLCJlbWFpbCI6InNlY3JldGFyaWFAZWR1YXJkb2dvbWVzLmFkdi5iciJ9.fjbbPDvwhwC6Ro8UPLG20Q4QEt3ZF1k09uaIB64XucT1wQDAwlxap4dxNG9S-zao_v33fGLvDO9efNsEoF3zy9D2j9c6_PTVbYVq66F8Gv_gF2QEeWDfDSe1IgcXkW5wTW3h4JuB2WdVocEaGvKaNlRX11FBo3FqJ212P8346nRfPUjoadfHUelPZ215FhMZiJLEJ9OKU_6NYJvucchvOtBZqyAU0tMda0VyOYy8VxN88i3F13br8KMjmlRWafwMnzyTHE1WZrZ0jXczcj5FuhLpKfSkvIbs8mZrNb3ccJJty6WOvsncvHA9UtISoT5Y1ioS27CJXI-XYfbqD_ZFFA",
    "eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3MzcyNzkxMDcsImlhdCI6MTczNzI1MDMwOCwiYXV0aF90aW1lIjoxNzM3MjUwMjg0LCJqdGkiOiJmM2FjZjkzYi0zNDVjLTRkODYtYTM1NS03NTI4NDhmOTNkMTMiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MTBlZTBhLWMzNzctNDhiMy05Y2VhLTNiZGQ5YzMzMjMzNSIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6ImE0YjMxMWFhLTE1ZTAtNGZhMC04MTAzLWVhNmI0ZTcyNjE2NSIsInNlc3Npb25fc3RhdGUiOiJmZjViM2QzMS00OTY0LTRmOGMtOTk3Mi1hMDgyZmRhMjNhODYiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJmZjViM2QzMS00OTY0LTRmOGMtOTk3Mi1hMDgyZmRhMjNhODYiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJHVUlMSEVSTUUgUEVSRUlSQSBESUFTIiwicHJlZmVycmVkX3VzZXJuYW1lIjoiMDQxNDg3ODIwNTUiLCJnaXZlbl9uYW1lIjoiR1VJTEhFUk1FIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkRJQVMiLCJlbWFpbCI6Imd1aXBkOTZ3YXJAZ21haWwuY29tIn0.04JeIRLF0Nmj9WuY5tY_TsCPJvp9Ad_0Qe9WFeyn0AsRECX2KwmgsG76FW1smZ5L-zrl-N5knT6Y6fpCWJNlbY-_fGJNCZ0ijOJGEh7Vz-sd2sytWzCY0jgpgHcjbXG5Wzj39kuFILlrGFJNd4375Iu_EhSw3qV886FxIAV5-M8u61bJWBz5XMG1CcgM9dZ5STdS1xgN9f4gEl-Hkz-UZjb4RuJBDZCbePt8EyWvQTZ5GBWCB81Tqnpdzdaq21qrSY3lpKo5vhpXSODXxgAbJwr1KPa4l0V9DD3kYhJykFCpZH_AzqoF1wyOnbHVnnlzRufuEcEmNgT2kRWpuv2fVQ",
    "eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3MzcyNzkxNDMsImlhdCI6MTczNzI1MDM0MywiYXV0aF90aW1lIjoxNzM3MjUwMzEwLCJqdGkiOiIxOWVlMGViYy04N2E5LTQ0NjItYjI1Ni01NDRhMzM1MzQwODQiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MzY4ZGQwLWMyM2YtNDMyOS1hZWIxLTkxY2ViMWQ4MzkzOCIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6ImY0NTUzMWJiLTRiYjctNGI5YS04NDY3LWNkZWRmZThiZjYzOCIsInNlc3Npb25fc3RhdGUiOiI0ZWRjZWU4OS01YWNhLTRhYTEtYWRiZC01MzFlN2JmNWNkNDciLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiI0ZWRjZWU4OS01YWNhLTRhYTEtYWRiZC01MzFlN2JmNWNkNDciLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJBUlRIVVIgREFMTU9MSU4gUk9IREUiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiIwMzk1ODc5ODA5OCIsImdpdmVuX25hbWUiOiJBUlRIVVIgREFMTU9MSU4iLCJmYW1pbHlfbmFtZSI6IlJPSERFIiwiZW1haWwiOiJhcnRodXJyb2hkZUBob3RtYWlsLmNvbSJ9.ISdxYCWA5xlL0LxAp2VybL-PlgCMdzZPqxS89Ot9-zGmHyD8ZPcduJUl8mQmk9ptBe16_Qhw2pL0op52HMcJqoIxioSLy8VxBmMMtwmQBLBc7zKRbjMYvcwzoYBHr_q1ACskmtuufJIHgVb_gtlmhrSSIPnaXN2ROqXsIl-4wj3jAdq8Sr9DzPueiIkiN1NIynP_5EV-AmW01y6CnseGjUKnaT5NYteRspfhXEWKuvBqUS85csHMAeXvAD2hz10kWq4E0l0kPmsPQ5NS2bEufPzIVw6QlK9JgC55_UJ24WD5J8CyJuXICZoAAZiIBmOTC46T_mjxPynkhe17zqycwA",
])
lock = threading.Lock()


def executar():
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    processos = capturar_processos_sheets()

    for processo in processos:
        token = next(BEARER_CODES)
        processar_processo(processo, token)

    salvar_informacoes_no_excel()


def processar_processo(processo, token):
    try:
        numero_processo = processo.get('numero_processo')
        print(f"Iniciando o processamento do processo {numero_processo}...")
        ultimo_movimento = processo.get('ultimo_movimento')

        dados = api_jusbr(numero_processo, token)
        dados_prontos, movimentos = dados_formatados(dados, numero_processo, ultimo_movimento)

        with lock:
            salvar_informacoes_no_json(dados_prontos, NOME_ARQUIVO_PARA_SALVAR)
            salvar_informacoes_no_json(movimentos, 'movimentos')

        print(f"Processamento do processo {numero_processo} finalizado.")
    except Exception as e:
        print(f"Erro no processamento do processo {numero_processo}: {e}")


def capturar_processos_sheets():
    processos_consulta_andamento = acessar_sheets('Procesos_Andamentos')
    df = pd.DataFrame(processos_consulta_andamento.get_all_records())
    processos = []

    for index, linha in df.iterrows():
        numero_processo = str(linha['Número do processo'])
        ultimo_movimento = linha['Data último movimento']
        processos.append({
            'numero_processo': numero_processo,
            'ultimo_movimento': datetime_datahora(ultimo_movimento, '%d/%m/%Y %H:%M:%S')
        })
    return processos


def api_jusbr(processo, token):
    processo_limpo = processo.replace('.', '').replace('-', '')
    url_base = f'https://portaldeservicos.pdpj.jus.br/api/v2/processos/{processo_limpo}'
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url_base, headers=headers)

    if response and response.status_code == 200:
        return response.json()


def dados_formatados(dados, numero_processo, ultimo_movimento):
    proximo_movimento_atual = None
    processo = dados[0].get('tramitacaoAtual')
    valor_acao = processo.get('valorAcao')
    grau = processo.get('grau', {}).get('nome')
    tribunal = processo.get('tribunal', {}).get('sigla')
    classe_principal = processo.get('classe', [{}])[0].get('descricao')

    data_de_distribuicao = datetime_datahora(processo.get('distribuicao', [{}])[0].get('dataHora', {}).split('.')[0])
    ultima_semana = datetime.now() - timedelta(days=7)
    movimentos = []

    for movimento in processo.get('movimentos'):
        movimento_data = datetime_datahora(movimento.get('dataHora').split('.')[0])
        ultimo_movimento = datetime_datahora(ultimo_movimento)

        if movimento_data > ultima_semana and (ultimo_movimento is None or movimento_data > ultimo_movimento):
            descricao_movimento = movimento.get('descricao')

            if not proximo_movimento_atual or proximo_movimento_atual < movimento_data:
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

    data_ultimo_movimento = proximo_movimento_atual if proximo_movimento_atual else ultimo_movimento

    informacoes_movimentos = {
        'Número do processo': numero_processo,
        'Data último movimento': data_ultimo_movimento.strftime('%d/%m/%Y %H:%M:%S') if data_ultimo_movimento else None
    }

    return informacoes, informacoes_movimentos


def salvar_informacoes_no_json(informacoes, arquivo):
    if os.path.exists(f'{arquivo}.json'):
        with open(f'{arquivo}.json', 'r', encoding='utf-8') as f:
            dados = json.load(f)
    else:
        dados = []

    dados.append(informacoes)

    with open(f'{arquivo}.json', 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


def salvar_informacoes_no_excel():
    aba = acessar_sheets('TESTE')
    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)

    dados_flat = [registro for grupo in dados for registro in grupo]
    df = pd.DataFrame(dados_flat)

    ultima_linha_preenchida = len(aba.get_all_values())
    linha_inicial = ultima_linha_preenchida + 1

    range_inicial = f'A{linha_inicial}'
    aba.update(range_inicial, df.values.tolist())

    if os.path.exists(f'{NOME_ARQUIVO_PARA_SALVAR}.json'):
        os.remove(f'{NOME_ARQUIVO_PARA_SALVAR}.json')

    print(f'o arquivo {NOME_ARQUIVO_PARA_SALVAR}.xlsx foi gerado com sucesso!')

    aba = acessar_sheets('Procesos_Andamentos')
    with open('movimentos.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)

    df = pd.DataFrame(dados)
    aba.update('A2', df.values.tolist())

    if os.path.exists('movimentos.json'):
        os.remove('movimentos.json')


def acessar_sheets(aba):
    url_apatir_do_d = '1AzdNQyAm_smi-t3tNwrMQHUiir42flWsMnJpdgVA6-k'
    google_cloud = gspread.service_account(filename='bancarioadvbox-e9fd143ea26d.json')
    sheet = google_cloud.open_by_key(url_apatir_do_d)
    return sheet.worksheet(aba)


def datetime_datahora(data, format='%Y-%m-%dT%H:%M:%S'):
    if data:
        if isinstance(data, datetime):
            return data
        return datetime.strptime(data, format)
    return None


executar()
import json
import locale
import os
import threading
import time
from datetime import datetime, timedelta
from itertools import cycle

import gspread
import pandas as pd
import requests

NOME_ARQUIVO_PARA_SALVAR = 'andamentos_processos'
BEARER_CODES = cycle([
    "eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3MzgxMzgyMzEsImlhdCI6MTczODEwOTQzMSwiYXV0aF90aW1lIjoxNzM4MTA5NDIzLCJqdGkiOiIzNjM2OGNiMC1jZDNiLTQyYWQtYTY4Ni00YTZmZDcwYTkyNGQiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MTBlZTBhLWMzNzctNDhiMy05Y2VhLTNiZGQ5YzMzMjMzNSIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6IjkzZGVjYzgwLWVmNzEtNGFkNC05MjMzLTRjZDFlYjI5MTNiYiIsInNlc3Npb25fc3RhdGUiOiJkMTlmYjQzOC1jMmYyLTQ5ZjYtOTA4MC03MTQzNGU3ZWU4ZmIiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJkMTlmYjQzOC1jMmYyLTQ5ZjYtOTA4MC03MTQzNGU3ZWU4ZmIiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJHVUlMSEVSTUUgUEVSRUlSQSBESUFTIiwicHJlZmVycmVkX3VzZXJuYW1lIjoiMDQxNDg3ODIwNTUiLCJnaXZlbl9uYW1lIjoiR1VJTEhFUk1FIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkRJQVMiLCJlbWFpbCI6Imd1aXBkOTZ3YXJAZ21haWwuY29tIn0.ufGGda4zmmBHvnBW-uhxz6fRTU4m1lVOdbbN9N0cph_-OBch9yr7TOMZF5dEP_VnPtlGh0aE8oyDbCIqVfNDdaC8dpzs8HzUZlD9YFhef81Zc7PkbNEDVgpYHD2tGJRKgFKcoJzawq7GvQtOdiYgNQdqHTHne31-URMU2ceVGsoKIJ12Wld6BnddJk_PbVyxicv04CnOQTKfhhujzArBG_-6MVgN_9OsG0e3lgHsjmix_LCtlZyqBy7OZqSQIGQ_E5lLqzIxV4FSyOGC8hZ18nbNvbOmwGn4Xo-KEHHnNYwQPUTG4QBpyZb7F6kiZTMYV59sIjdlNCC5mbpAM5opTg",
    "eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3MzgxMzgyMzEsImlhdCI6MTczODEwOTQzMSwiYXV0aF90aW1lIjoxNzM4MTA5NDIzLCJqdGkiOiIzNjM2OGNiMC1jZDNiLTQyYWQtYTY4Ni00YTZmZDcwYTkyNGQiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MTBlZTBhLWMzNzctNDhiMy05Y2VhLTNiZGQ5YzMzMjMzNSIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6IjkzZGVjYzgwLWVmNzEtNGFkNC05MjMzLTRjZDFlYjI5MTNiYiIsInNlc3Npb25fc3RhdGUiOiJkMTlmYjQzOC1jMmYyLTQ5ZjYtOTA4MC03MTQzNGU3ZWU4ZmIiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJkMTlmYjQzOC1jMmYyLTQ5ZjYtOTA4MC03MTQzNGU3ZWU4ZmIiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJHVUlMSEVSTUUgUEVSRUlSQSBESUFTIiwicHJlZmVycmVkX3VzZXJuYW1lIjoiMDQxNDg3ODIwNTUiLCJnaXZlbl9uYW1lIjoiR1VJTEhFUk1FIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkRJQVMiLCJlbWFpbCI6Imd1aXBkOTZ3YXJAZ21haWwuY29tIn0.ufGGda4zmmBHvnBW-uhxz6fRTU4m1lVOdbbN9N0cph_-OBch9yr7TOMZF5dEP_VnPtlGh0aE8oyDbCIqVfNDdaC8dpzs8HzUZlD9YFhef81Zc7PkbNEDVgpYHD2tGJRKgFKcoJzawq7GvQtOdiYgNQdqHTHne31-URMU2ceVGsoKIJ12Wld6BnddJk_PbVyxicv04CnOQTKfhhujzArBG_-6MVgN_9OsG0e3lgHsjmix_LCtlZyqBy7OZqSQIGQ_E5lLqzIxV4FSyOGC8hZ18nbNvbOmwGn4Xo-KEHHnNYwQPUTG4QBpyZb7F6kiZTMYV59sIjdlNCC5mbpAM5opTg"
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
        tipo_processo = processo.get('tipo_processo')
        ultimo_movimento = processo.get('ultimo_movimento')
        print(f"Iniciando o processamento do processo {numero_processo}...")

        dados = api_jusbr(numero_processo, token)
        if dados:
            dados_prontos, movimentos = dados_formatados(dados, numero_processo, ultimo_movimento, tipo_processo)

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
        tipo_processo = linha['Tipo']
        processos.append({
            'numero_processo': numero_processo,
            'ultimo_movimento': datetime_datahora(ultimo_movimento, '%d/%m/%Y %H:%M:%S'),
            'tipo_processo': tipo_processo
        })
    return processos


def api_jusbr(processo, token):
    processo_limpo = processo.replace('.', '').replace('-', '')
    url_base = f'https://portaldeservicos.pdpj.jus.br/api/v2/processos/{processo_limpo}'
    headers = {'Authorization': f'Bearer {token}'}

    for _ in range(3):
        try:
            response = requests.get(url_base, headers=headers, timeout=10)
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Erro na requisição: {response.status_code} no processo {processo}")
        except requests.exceptions.Timeout:
            print(f"Timeout na requisição para o processo {processo}. Tentando novamente...")
            print(f"Erro no processo {processo}")
            time.sleep(1)

    return None  # Retorna None caso falhe após 3 tentativas


def dados_formatados(dados, numero_processo, ultimo_movimento, tipo_processo):
    proximo_movimento_atual = None
    processo = dados[0].get('tramitacaoAtual')
    valor_acao = processo.get('valorAcao')
    grau = processo.get('grau', {}).get('nome')
    tribunal = processo.get('tribunal', {}).get('sigla')
    classe_principal = processo.get('classe', [{}])[0].get('descricao')
    descricao_sentenca = None
    data_sentenca = None

    data_de_distribuicao = datetime_datahora(processo.get('distribuicao', [{}])[0].get('dataHora', {}).split('.')[0])
    ultima_semana = datetime.now() - timedelta(days=7)
    movimentos = []

    for movimento in processo.get('movimentos'):

        movimento_data = datetime_datahora(movimento.get('dataHora').split('.')[0])
        ultimo_movimento = datetime_datahora(ultimo_movimento)
        descricao_movimento = movimento.get('descricao')

        if movimento_data > ultima_semana and (ultimo_movimento is None or movimento_data > ultimo_movimento):

            if not proximo_movimento_atual or proximo_movimento_atual < movimento_data:
                proximo_movimento_atual = movimento_data

            movimento_data = movimento_data.strftime('%d/%m/%Y')
            movimentos.append([movimento_data, descricao_movimento])

        sentenca = any(sentenca in descricao_movimento.lower() for sentenca in SENTENCAS)

        if sentenca and not descricao_sentenca:
            data_sentenca = movimento_data.strftime('%d/%m/%Y')
            descricao_sentenca = movimento.get('descricao')


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
            'Polo ativo': nome_ativo.title() if nome_ativo else "Nome não informado",
            'Polo passivo': nome_passivo.title() if nome_passivo else "Nome não informado",
            'Tipo': tipo_processo

        })

    data_ultimo_movimento = proximo_movimento_atual if proximo_movimento_atual else ultimo_movimento

    informacoes_movimentos = {
        'Número do processo': numero_processo,
        'Data último movimento': data_ultimo_movimento.strftime('%d/%m/%Y %H:%M:%S') if data_ultimo_movimento else None,
        'Sentença': descricao_sentenca,
        'Data da Sentença': data_sentenca
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

SENTENCAS = [
    "trânsito em julgado",
    "extinto",
    "arquivado",
    "baixado",
    "procedente",
    "procedência",
    "improcedente",
    "improcedência",
    "transitado"
]
executar()
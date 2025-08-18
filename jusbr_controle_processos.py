import json
import locale
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from itertools import cycle

import gspread
import pandas as pd
import requests

NOME_ARQUIVO_PARA_SALVAR = 'andamentos_processos'
BEARER_CODES = cycle(['eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3NTI1NTAyNzEsImlhdCI6MTc1MjUyMTQ3MSwiYXV0aF90aW1lIjoxNzUyNTIxMDUzLCJqdGkiOiI4YjA0ODJlNS1jNDFlLTRiMmQtYTE4Ny04YjRmYWE4M2Y2YjkiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6IjhkMGMzYmNjLTNkOWItNGZlMy04ZThjLWFhN2M0Mzk5NGEwYiIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6IjdkNWQ0NDViLTFmM2MtNDBkMy05MTQ2LWUyZGFkM2Q0M2RmZCIsInNlc3Npb25fc3RhdGUiOiIzMjg3M2Q4Zi1jMzgxLTQ2ZTEtYWE1Mi05NjIwY2Y5NjMwODEiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiIzMjg3M2Q4Zi1jMzgxLTQ2ZTEtYWE1Mi05NjIwY2Y5NjMwODEiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJFRFVBUkRPIFBFUkVJUkEgR09NRVMiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiI3Njc4NzA0NDAyMCIsImdpdmVuX25hbWUiOiJFRFVBUkRPIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkdPTUVTIiwiY29ycG9yYXRpdm8iOlt7InNlcV91c3VhcmlvIjo1MzQ3MjA5LCJub21fdXN1YXJpbyI6IkVEVUFSRE8gUEVSRUlSQSBHT01FUyIsIm51bV9jcGYiOiI3Njc4NzA0NDAyMCIsInNpZ190aXBvX2NhcmdvIjoiQURWIiwiZmxnX2F0aXZvIjoiUyIsInNlcV9zaXN0ZW1hIjowLCJzZXFfcGVyZmlsIjowLCJkc2Nfb3JnYW8iOiJPQUIiLCJzZXFfdHJpYnVuYWxfcGFpIjowLCJkc2NfZW1haWwiOiJzZWNyZXRhcmlhQGVkdWFyZG9nb21lcy5hZHYuYnIiLCJzZXFfb3JnYW9fZXh0ZXJubyI6MCwiZHNjX29yZ2FvX2V4dGVybm8iOiJPQUIiLCJvYWIiOiJSUzkxNjMxIn1dLCJlbWFpbCI6InNlY3JldGFyaWFAZWR1YXJkb2dvbWVzLmFkdi5iciJ9.kYdvAZXzjmkx_zrGEjDtlvstxba9QBdt6AIHOVvc6ogi2nFlX9CkzQHt5Q5qboudqfCik28GqRIGN7u-x1yLC0bHn53gDdcw3B92cnxOrsfBftIB5LqIa7XUsJT41w8D9u4DOE6niTiYF0cVzY4gnDaYo_N_G4AT85ux32UytFhMsBHlQM3V8Iy8UYTU5ylmM_TvLohQJ-3Zhs6029ZYlX80mYYcsgyOTKF8i4owktQV__WXufbeOtz3kMYN78jcPONG8X2gpfmlqjJYCrn1o2DXHfKvmY7JLIaEN3yFXBc4khC1AVhnp8R6UbbDE9Ar705SGoTw7OJcd9F8FA95Nw',''])

lock = threading.Lock()


def executar():
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    processos = capturar_processos_sheets()

    with ThreadPoolExecutor(max_workers=1) as executor:
        executor.map(processar_processo, processos)

    salvar_informacoes_no_excel_2()

def processar_processo(processo):
    try:
        token = next(BEARER_CODES)
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
    df = pd.read_excel("processos_jusbr.xlsx", sheet_name="Plan1")
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
    # processo_limpo = '00000011220174036000'
    url_base = f'https://portaldeservicos.pdpj.jus.br/api/v2/processos/{processo_limpo}'
    headers = {'Authorization': f'Bearer {token}'}

    for _ in range(1):
        try:
            response = requests.get(url_base, headers=headers)
            if response.status_code == 200:
                print(f"foi")
                return response.json()
            else:
                print(f"Erro na requisição: {response.status_code} no processo {processo}")
        except requests.exceptions.Timeout:
            print(f"Timeout na requisição para o processo {processo}. Tentando novamente...")
            print(f"Erro no processo {processo}")
            time.sleep(1)

    return None


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

        if not sentenca:
            sentenca = any(sentenca in descricao_movimento.lower() for sentenca in SENTENCAS_SECUNDARIAS)

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


def salvar_informacoes_no_excel_2():
    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)

    dados_flat = [registro for grupo in dados for registro in grupo]
    df = pd.DataFrame(dados_flat)
    df.to_excel(f'{NOME_ARQUIVO_PARA_SALVAR}.xlsx', index=False, engine='openpyxl', sheet_name="Processos")

    with open('movimentos.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)

    df2 = pd.DataFrame(dados)
    df2.to_excel(f'movimentos.xlsx', index=False, engine='openpyxl', sheet_name="Processos")

    print(f'o arquivo {NOME_ARQUIVO_PARA_SALVAR}.xlsx foi gerado com sucesso!')


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
    "audiência de instrução"
]

SENTENCAS_SECUNDARIAS = [
    "Sentença",
    "Sentenciado",
    "Julgamento "
]


# salvar_informacoes_no_excel_2()
executar()
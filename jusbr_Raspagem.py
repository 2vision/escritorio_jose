import json
import locale
import os
import time
from datetime import datetime
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor

NOME_ARQUIVO_PARA_SALVAR = 'NOVO'

CNPJS = [
    "00.000.000/0001-91",
    "00.360.305/0001-04",
    "60.701.190/0001-04",
    "60.746.948/0001-12",
    "90.400.888/0001-42",
    "58.160.789/0001-28",
    "30.306.294/0001-45",
    "59.588.111/0001-03",
    "17.184.037/0001-10",
    "59.285.411/0001-13",
    "18.236.120/0001-58",
    "00.416.968/0001-01",
    "31.872.495/0001-72",
    "20.855.875/0001-82",
    "92.894.922/0001-08",
    "08.561.701/0001-01",
    "61.186.680/0001-74",
    "30.723.886/0001-62",
]


BEARER_CODE = 'eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3NTA1NTEyOTEsImlhdCI6MTc1MDUyMjQ5MiwiYXV0aF90aW1lIjoxNzUwNTIyNDgzLCJqdGkiOiJjZGY4MDA2NC0yOTY2LTQ3NDgtODQ2Ni1lZTcxMjE2YmNmNTUiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MTBlZTBhLWMzNzctNDhiMy05Y2VhLTNiZGQ5YzMzMjMzNSIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6ImFkNDU3MmI1LWMzMjMtNDk5My1hYzYxLWI3ZjJmN2Y5YmM4YSIsInNlc3Npb25fc3RhdGUiOiJiYzQ5OThjOC1kYjdjLTQzMzktOTY1NS0xYzBlOGNlMDYyN2UiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJiYzQ5OThjOC1kYjdjLTQzMzktOTY1NS0xYzBlOGNlMDYyN2UiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJHVUlMSEVSTUUgUEVSRUlSQSBESUFTIiwicHJlZmVycmVkX3VzZXJuYW1lIjoiMDQxNDg3ODIwNTUiLCJnaXZlbl9uYW1lIjoiR1VJTEhFUk1FIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkRJQVMiLCJlbWFpbCI6Imd1aXBkOTZ3YXJAZ21haWwuY29tIn0.pkPUwOJJvsxk-mS-jityQo_4DWDSTzIsvU2_JCGWCtn6bUwQ8Sg1WtMbPSj6S7bTzWJk0Fpn4Va96sIzOaDgrP2IHW16jHn2AHrLfaamSCW1GiDld9Gip1fKG8gJ93c6w9ZpjcQwEaWECBRfQuWg1K8s7ud25SBFkQKFuWCfe_bNHQ_a4glVprPMtmIGEZrAaVsyTn4yLQFeLKantc2FDIaKIdnZYo1iHkWy-UuW9Tmsan5Y5nbJ-LoyU2myb77GO0hpTC44-9HMYUWITCicRIg9S6UvEmWlWfaP9RIAVz6KCDmKO1uMC2Wk0_QGGR-LMoIGtQnKUqxs7u-BfMeWZA'
DATA_INICIAL = '01/06/2025'
VALOR_MINIMO = '30000,00'
TRIBUNAIS = ['TJRS']
CLASSES = ['Execução de Título Extrajudicial', 'Monitória', 'Procedimento Comum Cível', 'Execução Fiscal', 'Busca e apreensão']

def processar_cnpj(cnpj, tribunal, bearer_code, codigos_das_classes, data_inicial, valor_minimo):
    print(f'CNPJ {cnpj} e o tribunal {tribunal} iniciaram a análise.')

    proxima_pagina = None
    total_de_processos = 0
    quantidade_analisada = 0
    dados_dos_processos = []

    tem_proxima_pagina = True
    max_paginas = 15
    contador_paginas = 0

    while tem_proxima_pagina and contador_paginas < max_paginas:
        dados_da_pagina = api_jusbr(bearer_code, cnpj, tribunal, proxima_pagina)

        if not dados_da_pagina:
            break

        proxima_pagina = f'{dados_da_pagina.get("searchAfter")[0]},{dados_da_pagina.get("searchAfter")[1]}'
        processos = dados_da_pagina.get('content', {})

        quantidade_analisada += len(processos)
        total_de_processos = dados_da_pagina.get('total')
        tem_proxima_pagina = quantidade_analisada < total_de_processos

        contador_paginas += 1  # Incrementa o contador de páginas

        for processo in processos:
            dados_do_processo = capturar_informacoes(processo, cnpj, codigos_das_classes, data_inicial, valor_minimo)
            if dados_do_processo:
                salvar_informacoes_no_json(dados_do_processo)
                dados_dos_processos.append(dados_do_processo)

    return len(dados_dos_processos)


def executar():
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

    bearer_code = BEARER_CODE.replace('bearer', '').replace('Bearer', '').strip()
    cnpjs = [cnpj.replace('-', '').replace('.', '').replace('/', '') for cnpj in CNPJS]

    data_inicial = datetime.strptime(DATA_INICIAL, '%d/%m/%Y')
    valor_minimo = float(VALOR_MINIMO.replace('R$', '').replace('.', '').replace(',', '.')) if VALOR_MINIMO else None
    codigos_das_classes = [lista_de_classes[classe] for classe in CLASSES] if CLASSES else None
    tribunais = [lista_de_tribunais[tribunal] for tribunal in TRIBUNAIS] if TRIBUNAIS else [None]

    total_capturado = 0

    # Usando ThreadPoolExecutor para multithreading
    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = []
        for cnpj in cnpjs:
            for tribunal in tribunais:
                futures.append(
                    executor.submit(processar_cnpj, cnpj, tribunal, bearer_code, codigos_das_classes, data_inicial,
                                    valor_minimo))

        for future in futures:
            total_capturado += future.result()

    print(f'Total capturado: {total_capturado}')
    return total_capturado


def api_jusbr(bearer_code, cnpj, codigo_do_tribunal=None, paginacao=None):
    url_base = 'https://portaldeservicos.pdpj.jus.br/api/v2/processos'
    headers = {'Authorization': f'Bearer {bearer_code}'}
    query_params = {'cpfCnpjParte': cnpj}

    if paginacao:
        query_params['searchAfter'] = paginacao
    if codigo_do_tribunal:
        query_params['tribunal'] = codigo_do_tribunal

    max_retries = 1

    for tentativa in range(1, max_retries + 1):
        try:
            response = requests.get(url_base, headers=headers, params=query_params)

            if response.status_code == 200:
                return response.json()

            if response.status_code in {500, 502, 503, 504}:  # Erros do servidor
                print(f"Tentativa {tentativa}/{max_retries} falhou com erro {response.status_code}. Retentando ...")
            else:
                print(f"Erro {response.status_code}: {response.text}")
                return None

        except requests.exceptions.RequestException as e:
            print(f"Erro de conexão: {e}. Tentativa {tentativa}/{max_retries}. Retentando ...")

    print(f"Falha após {max_retries} tentativas para {cnpj} e {codigo_do_tribunal}.")
    return None

def capturar_informacoes(processo, cnpj, codigos_das_classes, data_inicial, valor_minimo):
    nome_polo_passivo, numero_documento = None, None
    documentos_iguais = []

    ultimo_tramites = processo.get('tramitacoes', [{}])[0]
    classes = ultimo_tramites.get('classe', [{}])
    partes = ultimo_tramites.get('partes', [{}])
    nome_polo_ativo = None
    nome_classe = None
    representante_advogado_polo_passivo = None

    for parte in partes:
        polo = parte.get('polo')
        if polo == 'ATIVO':
            nome_polo_ativo = parte.get('nome')
            documentos_principais = parte.get('documentosPrincipais')
            if documentos_principais:
                for documento in documentos_principais:
                    numero_documento = documento.get('numero')

                    if numero_documento == cnpj:
                        documentos_iguais.append(numero_documento)

    if codigos_das_classes:
        a_classe_tem_o_codigo = [classe.get('codigo') for classe in classes if
                                 classe.get('codigo') in codigos_das_classes]

        if a_classe_tem_o_codigo:
            nome_classe = next((nome for nome, num in lista_de_classes.items() if num == a_classe_tem_o_codigo[0]),
                               None)

    else:
        a_classe_tem_o_codigo = True

    if a_classe_tem_o_codigo and documentos_iguais:
        numero_processo = processo.get('numeroProcesso')
        estado_tribunal = processo.get('siglaTribunal')

        valor_acao = ultimo_tramites.get('valorAcao')
        ultima_distribuicao = ultimo_tramites.get('dataHoraUltimaDistribuicao')
        data_distribuicao = datetime.strptime(ultima_distribuicao.split('.')[0],
                                              '%Y-%m-%dT%H:%M:%S') if ultima_distribuicao else datetime(1500, 1, 1)

        for parte in partes:
            polo = parte.get('polo')
            if polo == 'PASSIVO':
                nome_polo_passivo = parte.get('nome')
                numero_documento = parte.get('documentosPrincipais', [{}])[0].get('numero')
                representante_advogado_polo_passivo = parte.get('representantes')

        verificacao_data = not data_inicial or data_inicial < data_distribuicao
        verificacao_valor = not valor_minimo or valor_minimo < valor_acao

        if verificacao_data and verificacao_valor and not representante_advogado_polo_passivo:
            return {
                'CPF/CNPJ': formatar_documento(numero_documento) if numero_documento else None,
                'Data da distribuição': data_distribuicao.strftime('%d/%m/%Y'),
                'Cliente': nome_polo_passivo.title() if nome_polo_passivo else 'Desconhecido',
                'Banco': nome_polo_ativo,
                'Nº do Processo': numero_processo,
                'Tribunal': estado_tribunal,
                'Valor Causa': locale.currency(valor_acao, grouping=True),
                'Classe Judicial': nome_classe,
            }


def formatar_documento(documento):
    if len(documento) == 11:
        return f"{documento[:3]}.{documento[3:6]}.{documento[6:9]}-{documento[9:]}"
    elif len(documento) == 14:
        return f"{documento[:2]}.{documento[2:5]}.{documento[5:8]}/{documento[8:12]}-{documento[12:]}"
    else:
        return None


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
    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)

    df = pd.DataFrame(dados)
    df.to_excel(f'{NOME_ARQUIVO_PARA_SALVAR}.xlsx', index=False, engine='openpyxl', sheet_name="Processos")

    if os.path.exists(f'{NOME_ARQUIVO_PARA_SALVAR}.json'):
        os.remove(f'{NOME_ARQUIVO_PARA_SALVAR}.json')

    print(f'o arquivo {NOME_ARQUIVO_PARA_SALVAR}.xlsx foi gerado com sucesso!')


lista_de_tribunais = {
    'TRF1': 401,
    'TRF2': 402,
    'TRF3': 403,
    'TRF4': 404,
    'TRF5': 405,
    'TRF6': 406,
    'TJDFT': 807,
    'TJAC': 801,
    'TJAL': 802,
    'TJAP': 803,
    'TJAM': 804,
    'TJBA': 805,
    'TJCE': 806,
    'TJES': 808,
    'TJGO': 809,
    'TJMA': 810,
    'TJMT': 811,
    'TJMS': 812,
    'TJMG': 813,
    'TJPA': 814,
    'TJPB': 815,
    'TJPR': 816,
    'TJPE': 817,
    'TJPI': 818,
    'TJRJ': 819,
    'TJRN': 820,
    'TJRS': 821,
    'TJRO': 822,
    'TJRR': 823,
    'TJSC': 824,
    'TJSP': 826,
    'TJSE': 825,
    'TJTO': 827,
}

lista_de_classes = {
    'Execução de Título Extrajudicial': 12154,
    'Execução Fiscal': 1116,
    'Monitória': 40,
    'Procedimento Comum Cível': 7,
    'Busca e apreensão': 20
}

executar()

salvar_informacoes_no_excel()
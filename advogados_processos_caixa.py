import json
import os
import time
from concurrent.futures import ThreadPoolExecutor
from itertools import cycle

import requests
from datetime import datetime
import locale
import pandas as pd

NOME_ARQUIVO_PARA_SALVAR = 'advogados_caixa'
BEARER_CODE = cycle(['eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3NDYyMzUwNjEsImlhdCI6MTc0NjIwNjI2MSwiYXV0aF90aW1lIjoxNzQ2MjA2MjUxLCJqdGkiOiIzZWNkYjYwZS01YWU5LTRjMzAtYTRmOS1jZDRhMjE5NGNhNDEiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MTBlZTBhLWMzNzctNDhiMy05Y2VhLTNiZGQ5YzMzMjMzNSIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6IjQ5ZWU0MTZkLWEyMzUtNGM3MC04N2IxLWI3MzFmYjI1Zjg0NyIsInNlc3Npb25fc3RhdGUiOiI1MjhhN2FmNC1kZWEwLTQzMmQtOGQ3OC1kMGM4ZTUyOWFlZmEiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiI1MjhhN2FmNC1kZWEwLTQzMmQtOGQ3OC1kMGM4ZTUyOWFlZmEiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJHVUlMSEVSTUUgUEVSRUlSQSBESUFTIiwicHJlZmVycmVkX3VzZXJuYW1lIjoiMDQxNDg3ODIwNTUiLCJnaXZlbl9uYW1lIjoiR1VJTEhFUk1FIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkRJQVMiLCJlbWFpbCI6Imd1aXBkOTZ3YXJAZ21haWwuY29tIn0.ljN1xfxP1HGhLNPHmSDP28MJ0D7_PPavCal0nqWQpYKxgeyvLXgO19e2ftV9jea0NCoz5Z2TJV8s2aoz4_fyY1WZUxJH4QZ4qaqUfLL7knX-T6_bYXL6mAFda0kqsve5Cre9xePc-9114rGXJ8gA6W5rSTJVT5wlSIu92fZxg0DTf3fEuqpuq7DS59yIYcvv-yKEGf4UiLrrqnthgNoMoNCbasKmCajHr-59rprlAYlGFFhcdZiW6SXikDJVmGCuDvzPUKigcgcFWI04Rf4ymMEOPc7VHLCbnWLkzXYUpvZDotYfqnUf7S9_NvydOl7kb5W-mhbp-ferK40QWO3-jQ'])

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

df_processos = pd.read_excel('caixa.xlsx')  # Lê o Excel
numeros_de_processos = df_processos['Processos'].dropna().astype(str).str.zfill(20).tolist()

def consultar_processo_por_numero(numero_processo):
    time.sleep(1)
    token = next(BEARER_CODE)
    url = f'https://portaldeservicos.pdpj.jus.br/api/v2/processos'
    headers = {'Authorization': f'Bearer {token}'}
    params = {'numeroProcesso': numero_processo}

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        dados = response.json()
        if dados.get("content"):
            return capturar_informacoes(dados["content"][0])
        else:
            print(f'⚠️ Processo não encontrado: {numero_processo}')
    else:
        print(f'❌ Erro na requisição para {numero_processo} - Status: {response.status_code}')

    return []

def capturar_informacoes(processo):
    tramitacoes = processo.get('tramitacoes', [{}])
    ultima_tramitacao = tramitacoes[0] if tramitacoes else {}

    numero_processo = processo.get('numeroProcesso')
    estado_tribunal = processo.get('siglaTribunal')
    valor_acao = ultima_tramitacao.get('valorAcao', 0.0)

    ultima_distribuicao = ultima_tramitacao.get('dataHoraUltimaDistribuicao')
    data_distribuicao = (
        datetime.strptime(ultima_distribuicao.split('.')[0], '%Y-%m-%dT%H:%M:%S')
        if ultima_distribuicao else None
    )

    partes = ultima_tramitacao.get('partes', [])

    linhas_processadas = []

    for parte in partes:
        if not 'CAIXA' in parte.get('nome'):
            nome_polo_ativo = parte.get('nome')
            documento = parte.get('documentosPrincipais', [{}])[0].get('numero')

            advogado_nome = None
            advogado_documento = None

            representantes = parte.get('representantes', [])
            for rep in representantes:
                if rep.get('tipoRepresentacao') == 'ADVOGADO':
                    advogado_nome = rep.get('nome')
                    cpf_adv = rep.get('cadastroReceitaFederal', [{}])[0].get('numero')
                    advogado_documento = formatar_documento(str(cpf_adv)) if cpf_adv else None
                    break

            linha = {
                'Nº do Processo': numero_processo,
                'Tribunal': estado_tribunal,
                'Valor Causa': locale.currency(valor_acao, grouping=True),
                'Data da distribuição': data_distribuicao.strftime('%d/%m/%Y') if data_distribuicao else None,
                'Ativo': nome_polo_ativo,
                'CPF/CNPJ': formatar_documento(documento) if documento else None,
                'Advogado': advogado_nome,
                'Documento Advogado': advogado_documento,
            }

            linhas_processadas.append(linha)

    return linhas_processadas

def formatar_documento(documento):
    if len(documento) == 11:
        return f"{documento[:3]}.{documento[3:6]}.{documento[6:9]}-{documento[9:]}"
    elif len(documento) == 14:
        return f"{documento[:2]}.{documento[2:5]}.{documento[5:8]}/{documento[8:12]}-{documento[12:]}"
    return documento

def salvar_em_excel(dados):
    df = pd.DataFrame(dados)
    df.to_excel(f'{NOME_ARQUIVO_PARA_SALVAR}.xlsx', index=False)
    print(f'✅ Arquivo {NOME_ARQUIVO_PARA_SALVAR}.xlsx salvo com sucesso!')


def processar_processo(numero):
    try:
        info_linhas = consultar_processo_por_numero(numero)
        return info_linhas if info_linhas else []
    except Exception as e:
        print(f"Erro ao processar processo {numero}: {e}")
        return []

# Multithread com map
resultados = []
with ThreadPoolExecutor(max_workers=2) as executor:
    for resultado in executor.map(processar_processo, numeros_de_processos):
        resultados.extend(resultado)

if resultados:
    salvar_em_excel(resultados)

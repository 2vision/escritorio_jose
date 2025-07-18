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
BEARER_CODE = cycle(['eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3NDY0MTY3MjAsImlhdCI6MTc0NjM4NzkyMCwiYXV0aF90aW1lIjoxNzQ2Mzg3OTEzLCJqdGkiOiJiNmQ3YzQ4ZC1mYzMwLTQyMDAtYWU1Ny00OGZlNjBhNTk2ZWIiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6ImE3MTBlZTBhLWMzNzctNDhiMy05Y2VhLTNiZGQ5YzMzMjMzNSIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6IjU3NTczZjAwLWY3OWYtNDdjOC1iNDU5LTBjN2NkYjQ5ODc5YiIsInNlc3Npb25fc3RhdGUiOiJiNzdjMGEwMS04YTM2LTQyZWYtODUwNy03MzEyNjdlODgyMzMiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiJiNzdjMGEwMS04YTM2LTQyZWYtODUwNy03MzEyNjdlODgyMzMiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJHVUlMSEVSTUUgUEVSRUlSQSBESUFTIiwicHJlZmVycmVkX3VzZXJuYW1lIjoiMDQxNDg3ODIwNTUiLCJnaXZlbl9uYW1lIjoiR1VJTEhFUk1FIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkRJQVMiLCJlbWFpbCI6Imd1aXBkOTZ3YXJAZ21haWwuY29tIn0.ZSEF15cRIl5OQedItvQnHnwig5pWdZoKKPWdzEv70nVlwAEKLg_UxBupd6p9ah_5QWPbnNuVIVI0BdW-3HcpRDYoUzexl0GRaqC8BVkW9bTWdBrRp9eMQjNcwQ-CerLzQ3cCqx_AFDQuSCxkQ_IcxvDjto7eeOmQIldkWK9JQgKzJX5UmmyuQ3My7thjs1vtR84WtwaRFk5SQJG_27U2B0pgagw44OnMb8CD5KW_SpoZxTZ3jaIhPRN4qP4fI5JS29uMSUI7WSNHJz2F3s9CFJA0YCCMD05Ce48dP0dTcvzKmd7gOQ8NEWsOwIH289gOL8ESxJ5lNsEFzcavj1BP5A'])

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

import json
import os

import pandas as pd
import requests
from anticaptchaofficial.recaptchav2proxyless import recaptchaV2Proxyless
from concurrent.futures import ThreadPoolExecutor

NOME_ARQUIVO_PARA_SALVAR = 'MEDICOS SP'
API_KEY_ANTICAPTCHA = "7d64a894e2aa8e3c3cb916d76772fb57"
SITE_URL = "https://guiamedico.cremesp.org.br/"
SITE_KEY = "6LfafawZAAAAABQBiis7_2dboz4yyfGtuQjJBObK"
API_URL = "https://api.cremesp.org.br/guia-medico/filtrar"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "content-type": "application/json",
    "priority": "u=1, i",
    "sec-ch-ua": "\"Chromium\";v=\"136\", \"Google Chrome\";v=\"136\", \"Not.A/Brand\";v=\"99\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "Referer": SITE_URL,
}

def resolver_captcha():
    solver = recaptchaV2Proxyless()
    solver.set_verbose(0)
    solver.set_key(API_KEY_ANTICAPTCHA)
    solver.set_website_url(SITE_URL)
    solver.set_website_key(SITE_KEY)
    solver.set_is_invisible(1)

    token = solver.solve_and_return_solution()
    if token == 0:
        raise Exception(f"Erro ao resolver CAPTCHA: {solver.error_code}")
    return token

def fazer_requisicao_pagina(args):
    index_inicio, tamanho_pagina = args
    try:
        captcha_token = resolver_captcha()
        payload = {
            "indexInicioPagina": index_inicio,
            "tamanhoPagina": tamanho_pagina,
            "crm": None,
            "nome": "",
            "situacao": "A",
            "cidade": "",
            "especialidades": [],
            "reCaptcha": {
                "token": captcha_token
            }
        }

        requests.options(API_URL, headers=HEADERS)
        response = requests.post(API_URL, json=payload, headers=HEADERS)
        response.raise_for_status()
        return response.json()

    except Exception as e:
        print(f"Erro na requisição página {index_inicio // tamanho_pagina}: {e}")
        return None

def coletar_todos_dados():
    primeira_resposta = fazer_requisicao_pagina((0, 50))
    if not primeira_resposta:
        print("Falha na requisição inicial.")
        return []

    total_paginas = primeira_resposta.get("totalPages", 0)
    tamanho_pagina = primeira_resposta.get("size", 50)

    print(f"Total de páginas: {total_paginas}")
    lista_args = [(pagina * tamanho_pagina, tamanho_pagina) for pagina in range(1, total_paginas)]

    with ThreadPoolExecutor(max_workers=1) as executor:
        resultados = executor.map(fazer_requisicao_pagina, lista_args)

    for resultado in resultados:
        if resultado and "content" in resultado:
            todos_dados = resultado["content"]
            salvar_informacoes_no_json(todos_dados, 'todos_medicos')


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
    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)

    df = pd.DataFrame(dados)
    df.to_excel(f'{NOME_ARQUIVO_PARA_SALVAR}.xlsx', index=False, engine='openpyxl', sheet_name="Processos")

    if os.path.exists(f'{NOME_ARQUIVO_PARA_SALVAR}.json'):
        os.remove(f'{NOME_ARQUIVO_PARA_SALVAR}.json')

    print(f'o arquivo {NOME_ARQUIVO_PARA_SALVAR}.xlsx foi gerado com sucesso!')


dados = coletar_todos_dados()
print(f"Total de registros coletados: {len(dados)}")
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

import pandas as pd
import requests
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from seleniumwire import webdriver

LOGIN = 'raissa.zp'
SENHA = 'Raissa1001'
PUBLICATION_TYPES = [1, 4]


def executar():
    bearer_token = acessar_legal_one()

    processos = []
    lista_ids = []
    for publication_type in PUBLICATION_TYPES:
        processos_publication_type, lista_ids = lista_de_processos(bearer_token, publication_type, lista_ids)

        if processos_publication_type:
            processos.extend(processos_publication_type)

    processos_alterados = []
    if processos:
        print('Começou a dar baixa nos processos!')
        total_processos = len(processos)

        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_processo = {
                executor.submit(alterar_processo, bearer_token, processo): processo for processo in processos
            }

            for i, future in enumerate(as_completed(future_to_processo), 1):
                processo = future.result()
                if processo:
                    processos_alterados.append(processo)

                if i % 100 == 0 or i == total_processos:
                    print(f'{i}/{total_processos} processos concluídos...')

    else:
        print('Nenhum processo foi encontrado!')

    if processos_alterados:
        gerar_excel(processos_alterados)

    else:
        print('Nenhum processo foi alterado!')


def acessar_legal_one():
    bearer_token = None

    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')

    navegador = webdriver.Chrome(options=options)
    wait = WebDriverWait(navegador, 10)

    try:
        url = 'https://signon.thomsonreuters.com/v2?productid=L1NJ&returnto=https%3A%2F%2Flogin.novajus.com.br%2FOnePass%2FLoginOnePass%2F&bhcp=1'
        navegador.get(url)

        wait.until(EC.presence_of_element_located((By.ID, 'Username'))).send_keys(LOGIN)
        navegador.find_element(By.ID, 'Password').send_keys(SENHA)
        navegador.find_element(By.ID, 'SignIn').click()

        for request in navegador.requests:
            if 'https://legalone-prod-webapp-eastus2-api.azure-api.net' in request.url:
                bearer_token = request.headers.get('authorization')

    finally:
        navegador.quit()

        if bearer_token:
            return bearer_token


def lista_de_processos(bearer_token, publication_type, lista_ids):
    processos = []
    page = 1

    today = datetime.now()
    start_date = today - timedelta(days=60)

    filter_start_date = start_date.strftime('%Y-%m-%dT00:00:00.000Z')
    filter_end_date = today.strftime('%Y-%m-%dT00:00:00.000Z')

    while True:
        url = 'https://legalone-prod-webapp-eastus2-api.azure-api.net/prod//webapi/api/internal/publications/SearchPublicationsPaginated/'

        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
            'authenticationmethod': 'ASYMMETRIC_JWT_TOKEN',
            'authorization': bearer_token,
            'content-type': 'application/json',
            'distribution': 'FirmsBrazil',
            'ocp-apim-subscription-key': 'b1159d90df8d45148b4f5721e2752efc',
            'sec-ch-ua': '\'Not(A:Brand\';v=\'99\', \'Google Chrome\';v=\'133\', \'Chromium\';v=\'133\'',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '\'Linux\'',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'tenancy': 'mtadvogados',
            'Referer': 'https://firm.legalone.com.br/',
            'Referrer-Policy': 'strict-origin-when-cross-origin'
        }

        payload = {
            'page': page,
            'count': 200,
            'sortDirection': 1,
            'publicationType': publication_type,
            'filterDateOption': 0,
            'filterStartDate': filter_start_date,
            'filterEndDate': filter_end_date,
            'filterReadOption': 0,
            'filterCourtOption': [],
            'filterTreatmentOption': [0],
            'filterFontTypeOption': 0,
            'filterSourceOption': [],
            'filterResponsibleAreaOption': [],
            'filterResponsibleUserOption': [12204],
            'filterRelationshipFilterType': 0,
            'filterRelationshipFilterLawsuitOption': [],
            'filterRelationshipFilterContactOption': [],
            'filterTypePublicationOption': -1,
            'litigationTypes': [],
            'filterTagsOption': [],
            'filterStatesOption': [],
            'filterCourtSystemsOption': [],
            'filterCredentialOwnersOption': [],
            'customDateFilterStartDate': None,
            'customDateFilterEndDate': None,
            'customDateFilterOption': 2,
            'status': None,
            'filterTypeDateOption': 0,
            'type': '[Publication] getAll'
        }

        publication_type_text = 'Publicação' if publication_type == 1 else 'Intimação'
        print(f'Capturando processos da página {page} da {publication_type_text}')

        response = requests.post(url, headers=headers, data=json.dumps(payload))
        page += 1

        if response.status_code == 200:
            publications = response.json().get('data', {}).get('publications', [])

            if len(publications) > 0:
                for publication in publications:
                    if publication.get('publicationId') not in lista_ids:
                        lista_ids.append(publication.get('publicationId'))

                        data = publication.get('publishDate')
                        processos.append({
                            'id': publication.get('publicationId'),
                            'processo': publication.get('mainLitigation'),
                            'data_publicacao': datetime.fromisoformat(data).strftime('%d/%m/%Y') if data else None,
                            'tribunal': publication.get('journalInitials'),
                        })

            else:
                return processos, lista_ids

        else:
            print(f'Status code: {response.status_code}')
            print(f'Error: {response.json().get("error", {}).get("message")}')
            break


def alterar_processo(bearer_token, processo):
    url = "https://legalone-prod-webapp-eastus2-api.azure-api.net/prod//webapi/api/internal/publications/SetPublicationTreatStatus/"

    headers = {
        "accept": "application/json, text/plain, */*",
        "accept-language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "authenticationmethod": "ASYMMETRIC_JWT_TOKEN",
        "authorization": bearer_token,
        "content-type": "application/json",
        "distribution": "FirmsBrazil",
        "ocp-apim-subscription-key": "b1159d90df8d45148b4f5721e2752efc",
        "sec-ch-ua": "\"Not(A:Brand\";v=\"99\", \"Google Chrome\";v=\"133\", \"Chromium\";v=\"133\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Linux\"",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "cross-site",
        "tenancy": "mtadvogados",
        "Referer": "https://firm.legalone.com.br/",
        "Referrer-Policy": "strict-origin-when-cross-origin"
    }

    data = {
        "publicationId": processo.get('id'),
        "treatStatus": 2,
        "type": "[Publication] setPublicationTreatedStatus"
    }

    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 200:
        success = response.json().get('success', {})

        if success:
            return processo

    else:
        print(f'Status code: {response.status_code}')
        print(f'Error: {response.json().get("error", {}).get("message")}')

    return None


def gerar_excel(processos):
    data_hoje = datetime.today().strftime('%d-%m-%Y')

    df = pd.DataFrame(processos)
    df.rename(columns={
        'id': 'ID',
        'processo': 'Processos',
        'data_publicacao': 'Data de Publicação',
        'tribunal': 'Tribunal',
    }, inplace=True)

    df.to_excel(f"processos_{data_hoje}.xlsx", index=False)

    print("Arquivo Excel criado com sucesso!")


executar()

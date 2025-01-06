import os
import threading
import warnings
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from anticaptchaofficial.imagecaptcha import *
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from seleniumwire import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager

LOGIN = '00166687073'
SENHA = '@Dkz299302'
CAPTCHA_API_KEY = 'a89345c962e2eba448e571a6d0143363'
ARQUIVO = 'Dados_TRF4'
LISTA_PARCEIROS = []


def executar():
    processos_pendentes = pegar_processos()
    print('Iniciando captura de token para autenticação...')
    bearer_token = pegar_bearer_token(processos_pendentes[0])
    print('Token capturado.')
    print('Iniciando captura das informações dos processos...')

    while processos_pendentes:
        falhas = []

        def executar_todos_processos(processo):
            try:
                executar_processo(processo, bearer_token)
            except:
                falhas.append(processo)

        with ThreadPoolExecutor(max_workers=5) as executor:
            executor.map(executar_todos_processos, processos_pendentes)

        processos_pendentes = falhas

    salvar_informacoes_no_excel()
    print('Todos os processos foram finalizados.')


def executar_processo(processo, bearer_token):
    print(f'Processando o processo número {processo}...')
    # Request para pegar o ID do processo
    response_processo = request_processo(processo)
    processo_id = response_processo[0].get('id')

    # Request para pegar o TokenDesafio e a Imagem do Captcha
    response_processo_id = request_processo_id(processo_id)

    imagem = response_processo_id.get('imagem')
    token_desafio = response_processo_id.get('tokenDesafio')

    # Request para pegar a TaskID do Captcha
    response_captcha_id = request_captcha_imagem(imagem)

    # Request para pegar a resposta do Captcha
    response_captcha_text = request_captcha_id(response_captcha_id.get('taskId'))

    resposta = response_captcha_text.get('solution').get('text')

    # Request para pegar as informações do processo
    informacoes_processo = request_processo_captcha(processo_id, token_desafio, resposta, bearer_token)
    informacoes = padronizar_informacoes(informacoes_processo)

    with lock:
        salvar_informacoes_no_json(informacoes)
    print(f'Informações do processo {processo} salvas com sucesso!')


def pegar_processos():
    planilha_dados = pd.read_excel('TRT.xlsx', sheet_name='Plan1')
    return list(planilha_dados.iloc[:, 1])


def pegar_bearer_token(processo):
    bearer_token = None
    warnings.simplefilter(action='ignore', category=FutureWarning)

    edge_options = webdriver.EdgeOptions()
    edge_options.set_capability('ms:edgeChromium', True)
    edge_options.add_experimental_option('useAutomationExtension', False)
    edge_options.add_argument('--disable-blink-features=AutomationControlled')
    edge_options.add_argument('--disable-sync')
    edge_options.add_argument('--disable-features=msEdgeEnableNurturingFramework')
    edge_options.add_argument('--disable-popup-blocking')
    edge_options.add_argument('--disable-infobars')
    edge_options.add_argument('--disable-extensions')
    edge_options.add_argument('--remote-allow-origins=*')
    edge_options.add_argument('--disable-http2')
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    prefs = {
        'user_experience_metrics': {
            'personalization_data_consent_enabled': True
        }
    }

    edge_options.add_experimental_option('prefs', prefs)

    service = Service(EdgeChromiumDriverManager().install())
    navegador = webdriver.Edge(service=service, options=edge_options)
    wait = WebDriverWait(navegador, 10)

    navegador.maximize_window()

    navegador.get('https://pje.trt4.jus.br/primeirograu/login.seam')

    wait.until(EC.presence_of_element_located((By.ID, 'username'))).send_keys(LOGIN)
    navegador.find_element(By.ID, 'password').send_keys(SENHA)
    navegador.find_element(By.ID, 'btnEntrar').click()

    navegador.get('https://pje.trt4.jus.br/consultaprocessual/')

    wait.until(EC.visibility_of_element_located((By.ID, 'nrProcessoInput'))).send_keys(processo)
    navegador.find_element(By.ID, 'btnPesquisar').click()

    imagemCaptcha = wait.until(EC.visibility_of_element_located((By.ID, 'imagemCaptcha')))
    imagem = imagemCaptcha.get_attribute('src').split(',')[-1]

    response_captcha_id = request_captcha_imagem(imagem)
    response_captcha_text = request_captcha_id(response_captcha_id.get('taskId'))
    navegador.find_element(By.ID, 'captchaInput').send_keys(response_captcha_text.get('solution').get('text'))
    navegador.find_element(By.ID, 'btnEnviar').click()

    for request in navegador.requests:
        if 'api/processos/' in request.url and 'dadosbasicos' not in request.url:
            bearer_token = request.headers.get('authorization')

    navegador.quit()

    if bearer_token:
        return bearer_token

    else:
        raise Exception(
            f'Erro ao capturar o bearer_token utilizando Selenium'
        )


def request_processo(processo):
    url = f'https://pje.trt4.jus.br/pje-consulta-api/api/processos/dadosbasicos/{processo}'
    headers = {
        'x-grau-instancia': '1',
        'user-agent': 'Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.6778.200 Mobile Safari/537.36'
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(
            f'Erro na API TRF4 que captura o ID do processo (1ª request): {response.status_code} - {response.text}'
        )


def request_processo_id(processo_id):
    url = f'https://pje.trt4.jus.br/pje-consulta-api/api/processos/{processo_id}'
    headers = {
        'user-agent': 'Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.6778.200 Mobile Safari/537.36'
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(
            f'Erro na API TRF4 que captura pega a imagem do Captcha (2ª request): {response.status_code} - {response.text}'
        )


def request_captcha_imagem(imagem):
    url = 'https://api.anti-captcha.com/createTask'
    json = {
        'clientKey': CAPTCHA_API_KEY,
        'task': {
            'type': 'ImageToTextTask',
            'body': imagem,
            'phrase': False,
            'case': False,
            'numeric': False,
            'math': 0,
            'minLength': 0,
            'maxLength': 0
        }
    }

    response = requests.post(url, json=json)

    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(
            f'Erro na API Captcha que captura a taskId (3ª request): {response.status_code} - {response.text}'
        )


def request_captcha_id(task_id):
    url = 'https://api.anti-captcha.com/getTaskResult'
    json = {
        'clientKey': CAPTCHA_API_KEY,
        'taskId': task_id,
    }

    while True:
        time.sleep(.5)
        response = requests.post(url, json=json)

        if response.status_code == 200:
            if response.json().get('status') != 'processing':
                return response.json()


def request_processo_captcha(processo_id, token, resposta, bearer_token):
    url = f'https://pje.trt4.jus.br/pje-consulta-api/api/processos/{processo_id}?tokenDesafio={token}&resposta={resposta}'
    headers = {
        'x-grau-instancia': '1',
        'authorization': f'Bearer {bearer_token}',
        'user-agent': 'Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.6778.200 Mobile Safari/537.36'
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()

    else:
        print('Não foi possivel validar o AccessToken do processo. O processo será reprocessado...')
        raise Exception(
            f'Erro na API TRF4 que captura as informações do processo (5ª request): {response.status_code} - {response.text}'
        )


def padronizar_informacoes(informacoes_processo):
    informacoes = []
    if informacoes_processo.get('poloPassivo'):
        for reclamado in informacoes_processo.get('poloPassivo'):
            nome_reclamado = reclamado.get('nome').strip()
            data = datetime.fromisoformat(informacoes_processo.get('distribuidoEm')).strftime('%d/%m/%Y')
            valor = f"{informacoes_processo.get('valorDaCausa')}".replace('.', ',')

            nao_e_parceiro = not any(item.lower() in nome_reclamado.lower() for item in LISTA_PARCEIROS)

            if not LISTA_PARCEIROS or nao_e_parceiro:
                informacoes.append({
                    'Reclamado': nome_reclamado,
                    'Número do Processo': informacoes_processo.get('numero'),
                    'Reclamante': informacoes_processo.get('poloAtivo')[0].get('nome').strip(),
                    'Orgão Julgador': informacoes_processo.get('orgaoJulgador').strip(),
                    'Data de Distribuição': data,
                    'Valor da Causa': valor,
                    'Assuntos': ', '.join(
                        [assunto.get('descricao').strip() for assunto in informacoes_processo.get('assuntos')]),
                    'CPF/CNPJ': reclamado.get('documento'),
                })

    else:
        mensagemErro = informacoes_processo.get('mensagemErro')
        if mensagemErro:
            print(f'Não foi possível obter as informações do processo: {mensagemErro}')

        else:
            print('Não foi possivel validar o Captcha do processo. O processo será reprocessado...')
            raise Exception('Não foi possivel validar o Captcha do processo.')

    return informacoes


def salvar_informacoes_no_json(informacoes):
    if os.path.exists(f'{ARQUIVO}.json'):
        with open(f'{ARQUIVO}.json', 'r') as f:
            dados = json.load(f)
    else:
        dados = []

    dados.append(informacoes)

    with open(f'{ARQUIVO}.json', 'w') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


def salvar_informacoes_no_excel():
    with open(f'{ARQUIVO}.json', 'r') as f:
        dados = json.load(f)

    dados_flat = []
    for grupo in dados:
        for registro in grupo:
            dados_flat.append(registro)

    df = pd.DataFrame(dados_flat)
    df.to_excel(f'{ARQUIVO}.xlsx', index=False, sheet_name="Processos")

    if os.path.exists(f'{ARQUIVO}.json'):
        os.remove(f'{ARQUIVO}.json')


lock = threading.Lock()
executar()
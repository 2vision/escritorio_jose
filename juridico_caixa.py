import json
import os
import re
import warnings
import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from ftfy import fix_text
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading


NOME_ARQUIVO_PARA_SALVAR = 'Processos Verificados'
MOVIMENTOS = [
    "ACÓRDÃO DESFAVORÁVEL À CAIXA",
    "ACÓRDÃO FAVORÁVEL À CAIXA",
    "ACÓRDÃO PARCIALMENTE FAVORÁVEL",
    "ACORDO FIRMADO",
    "ARQUIVAMENTO",
    "AUDIÊNCIA DE CONC. OU INAUGURAL - ADV. CRED.",
    "AUDIÊNCIA DE INSTRUÇÃO E JULG. - ADV. CRED.",
    "AUDIÊNCIA DE JUSTIFICAÇÃO",
    "AUDIÊNCIA INSTITUCIONAL",
    "DECISÃO EMBARGOS DECLARAT.PARCIALM. ACOLHIDOS",
    "DECISÃO EMBARGOS DECLARATÓRIOS ACOLHIDOS",
    "DECISÃO EMBARGOS DECLARATÓRIOS REJEITADOS",
    "DECISÃO HOMOLOGATÓRIA DE DESISTÊNCIA",
    "DECISÃO HOMOLOGATÓRIA DE RENÚNCIA",
    "DECISÃO IMPOSIÇÃO SEGREDO DE JUSTIÇA",
    "DECISÃO INTERLOCUTÓRIA DECLÍNIO COMPETÊNCIA",
    "DECISÃO INTERLOCUTÓRIA DESFAVORÁVEL À CAIXA",
    "DECISÃO INTERLOCUTÓRIA FAVORÁVEL À CAIXA",
    "DECISÃO INTERLOCUTÓRIA PARCIALMENTE FAVORÁVEL",
    "DESISTÊNCIA DA AÇÃO PELA CAIXA PETIÇÃO",
    "DESISTÊNCIA DE RECURSO PELA CAIXA PETIÇÃO",
    "DESISTÊNCIA/RENÚNCIA PELA PARTE CONTRÁRIA",
    "DESPACHO DE MERO EXPEDIENTE INTIMAÇÃO",
    "DISTRIBUIÇÃO",
    "EMBARGOS À ADJUDICAÇÃO/ARREMATAÇÃO IMPUGNAÇÃO",
    "EMBARGOS À ADJUDICAÇÃO/ARREMATAÇÃO INTERPOSTO",
    "EMBARGOS À EXECUÇÃO/MONITÓRIA IMPUGNAÇÃO",
    "EMBARGOS À EXECUÇÃO/MONITÓRIA INTERPOSTOS",
    "EMBARGOS DE DECLARAÇÃO IMPUGNAÇÃO",
    "EMBARGOS DE DECLARAÇÃO INTERPOSTOS",
    "EMBARGOS DE TERCEIRO",
    "EMBARGOS INFRINGENTES CONTRARRAZÕES",
    "EMBARGOS INFRINGENTES INTERPOSTOS",
    "EXCEÇÃO DE IMPEDIMENTO INTERPOSTA",
    "EXCEÇÃO DE INCOMPETÊNCIA IMPUGNAÇÃO",
    "EXCEÇÃO DE INCOMPETÊNCIA INTERPOSTA",
    "EXCEÇÃO DE PRÉ-EXECUTIVIDADE IMPUGNAÇÃO",
    "EXCEÇÃO DE PRÉ-EXECUTIVIDADE INTERPOSTA",
    "EXCEÇÃO DE SUSPEIÇÃO",
    "EXECUÇÃO DE SENTENÇA IMPUGNAÇÃO",
    "EXECUÇÃO DE SENTENÇA PETIÇÃO",
    "EXECUÇÃO PROVISÓRIA INTIMAÇÃO",
    "EXTINCAO - AUSENCIA INTERESSE CAIXA/EMGEA",
    "FALENCIA - DECRETACAO OU CONVOLACAO",
    "IMPUGNAÇÃO À ASSISTÊNCIA JUDICIÁRIA GRATUITA",
    "IMPUGNAÇÃO AO VALOR DA CAUSA",
    "IMPUGNAÇÃO AO VALOR DA CAUSA RÉPLICA",
    "IMPUGNAÇÃO AOS CÁLCULOS",
    "INTIMAÇÃO DA PARTE CONTRÁRIA/LITISCONSORCIAL",
    "INTIMAÇÃO DE AUDIÊNCIA",
    "INTIMAÇÃO P/ LEVANTAMENTO DE ALVARÁ P/ CAIXA",
    "INTIMAÇÃO P/ MANIFESTAÇÃO PROCESSUAL- OUTRAS",
    "INTIMAÇÃO PARA CONTRARRAZÕES/RESPOSTA/IMPUG",
    "INTIMAÇÃO PAUTA DE JULGAMENTO",
    "INTIMAÇÃO PRODUÇÃO DE PROVAS",
    "INTIMAÇÃO RECURSO ADMITIDO",
    "INTIMAÇÃO RECURSO INADMITIDO",
    "INTIMAÇÃO REMESSA EX OFFICIO",
    "INTIMAÇÃO RESULTADO DE JULGAMENTO",
    "INTIMAÇÃO SUBSTITUIÇÃO DE PARTES",
    "INTIMAÇÃO SUSPENSÃO DO PROCESSO",
    "INTIMAÇÃO VISTA À(S) PARTE(S)",
    "INTIMAÇÃO/CITAÇÃO PARA EXECUÇÃO DE JULGADO",
    "LIMINAR - COMUNICACAO A AREA GESTORA",
    "LIMINAR CONCEDIDA",
    "LIMINAR NEGADA/REVOGADA/CASSADA",
    "LIMINAR PARCIALMENTE CONCEDIDA",
    "MANDADO DE IMISSÃO DE POSSE",
    "MANDADO DE SEGURANCA INTERPOSTO",
    "MANIFESTAÇÃO PROCESSUAL - OUTRAS",
    "NOTIFICAÇÃO EXTRAJUDICIAL",
    "PEDIDO DE ANTECIP. TUTELA/LIMINAR IMPUGNAÇÃO",
    "PRODUÇÃO DE PROVAS PETIÇÃO",
    "RAZÕES FINAIS OU MEMORIAIS PETIÇÃO",
    "RECURSO ADESIVO CONTRARRAZÕES",
    "RECURSO ADESIVO INTERPOSTO",
    "RECURSO ADM. DEFERIDO",
    "RECURSO ADM. INDEFERIDO",
    "RECURSO ADM. PARCIALMENTE DEFERIDO",
    "RECURSO ADMITIDO - APENAS EFEITO DEVOLUTIVO",
    "RECURSO DA CAIXA NÃO CONHECIDO",
    "RECURSO DE APELACAO- RESPOSTA",
    "RECURSO DE RECONSIDERAÇÃO",
    "RECURSO DE REVISTA INTERPOSTO",
    "RECURSO EM SENTIDO ESTRITO INTERPOSTO",
    "RECURSO ESPECIAL REPETITIVO",
    "RECURSO EXTRAORDINÁRIO - REPERCUSSÃO GERAL",
    "RECURSO EXTRAORDINÁRIO CONTRARRAZÕES",
    "RECURSO EXTRAORDINÁRIO INTERPOSTO",
    "RECURSO INOMINADO JEF INTERPOSTO",
    "RECURSO INOMINADO JEF RESPOSTA",
    "RECURSO MEDIDA CAUTELAR JEF",
    "RECURSO ORDINÁRIO CONTRARRAZÕES",
    "RECURSO ORDINÁRIO INTERPOSTO",
    "RECURSO PARCIALMENTE ADMITIDO",
    "RECURSO-CONTRARRAZÕES PTE CONTRÁRIA-LITISCONS",
    "SENTENÇA DESFAVORÁVEL À CAIXA",
    "SENTENÇA EXTINÇÃO POR CUMPRIMENTO OBRIGAÇÃO",
    "SENTENÇA EXTINÇÃO SEM RESOLUÇÃO DO MÉRITO",
    "SENTENÇA FAVORÁVEL À CAIXA",
    "SENTENÇA HOMOLOGATÓRIA DE ACORDO",
    "SENTENÇA HOMOLOGATÓRIA NA EXECUÇÃO",
    "SENTENÇA INCOMPETÊNCIA DE FORO",
    "SENTENÇA PARCIALMENTE FAVORÁVEL",
    "TRÂNSITO JULGADO DECISÃO DESFAVORÁVEL À CAIXA",
    "TRÂNSITO JULGADO DECISÃO FAVORÁVEL À CAIXA",
    "TRÂNSITO JULGADO DECISÃO PARCIALM. FAVORÁVEL",
    "TUTELA ANTECIPADA CONCEDIDA",
    "TUTELA ANTECIPADA NEGADA/REVOGADA/CASSADA",
    "TUTELA ANTECIPADA PARCIALMENTE CONCEDIDA",
    "A CLASSIFICAR"
]

def dados_planilha():
    pasta = 'Verificar Juridico Caixa'
    arquivos_no_diretorio = os.listdir(pasta)
    processos = []

    for arquivo in arquivos_no_diretorio:
        if arquivo.endswith('.xlsx'):
            caminho_arquivo = os.path.join(pasta, arquivo)
            df = pd.read_excel(caminho_arquivo)
            for _, linha in df.iterrows():
                numero_processo = str(linha['Processos'])
                data_publicacao_raw = linha['Data de Publicação']
                id = linha['ID']
                tribunal = linha['Tribunal']
                data_publicacao = pd.to_datetime(data_publicacao_raw, dayfirst=True).strftime('%d/%m/%Y')

                processos.append({
                    'numero_processo': numero_processo.replace('.', '').replace('-', ''),
                    'data_publicacao': data_publicacao,
                    'id': id,
                    'tribunal': tribunal
                })

    return processos


def pegar_cookies():
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
    navegador = webdriver.Edge(service=service)

    navegador.get("https://www.juridico.caixa.gov.br/")

    navegador.find_element(By.ID, 'sMatricula').send_keys('i003835')
    navegador.find_element(By.ID, 'sSenha').send_keys('df0883')

    WebDriverWait(navegador, 50).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#textoBuscaSijur"))
    )

    cookies_lista = navegador.get_cookies()
    cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies_lista}
    cookies = '; '.join(f'{k}={v}' for k, v in cookies_dict.items())

    return cookies


def consulta_numero_expediente(cookies, processo):
    url = "https://www.juridico.caixa.gov.br/?pg=busca"

    numero_do_processo = processo.get('numero_processo')

    headers = {
        'accept': 'text/html, */*; q=0.01',
        'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36',
        'x-requested-with': 'XMLHttpRequest',
        'Cookie': cookies
    }

    payload = f"por=processo&val={numero_do_processo}&bSomenteAtivos=false&bBuscaExata=true"
    response = requests.request("POST", url, headers=headers, data=payload)

    soup = BeautifulSoup(response.text, 'html.parser')

    texto_completo = soup.get_text()

    expedientes = re.findall(r'\d{2}\.\d{3}\.\d{5}/\d{4}(?:-\d{3})?', texto_completo)

    areas_judiciais_raw = re.findall(
        r'<td class="center">(.*?)</td>\s*<td class="center">(.*?)</td>\s*<td class="center">(.*?)</td>', response.text)
    areas_judiciais = [fix_text(re.sub(r'<.*?>', '', area[2]).strip()) for area in areas_judiciais_raw]

    expedientes = list(set(expedientes))

    return expedientes, areas_judiciais


def consulta_movimentos(cookies, numero_expediente, processo):
    url = f"https://www.juridico.caixa.gov.br/?pg=Expediente_movimentos&expediente={numero_expediente}&p_servidor=BDD7877WN001.corp.caixa.gov.br"
    data_planilha = datetime.strptime(processo.get('data_publicacao'), "%d/%m/%Y").date()
    data_dia_seguinte = data_planilha + timedelta(days=1)

    headers = {
        'accept': 'text/html, */*; q=0.01',
        'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36',
        'x-requested-with': 'XMLHttpRequest',
        'Cookie': cookies
    }

    response = requests.request("POST", url, headers=headers)

    soup = BeautifulSoup(response.text, 'html.parser')

    # Lista para armazenar os dicionários
    movimentos = []

    # Iterar sobre as linhas da tabela
    linhas = soup.select('#dadosMovimentos tbody tr')

    for linha in linhas:
        codigo = linha.select_one('.iNuFase')
        data = linha.select_one('.dDtFase')
        resumo = linha.select_one('.detalhes > div')
        texto_rtf = linha.select_one('.textoRtf')

        formatos = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y"]
        for fmt in formatos:
            try:

                data_formatada = datetime.strptime(data.text.strip(), fmt).date()

            except ValueError:
                continue

        movimentos.append({
            'codigo_movimento': fix_text(codigo.text.strip()) if codigo else '',
            'data': data_formatada,
            'descricao_resumida': fix_text(resumo.text.strip()) if resumo else '',
            'descricao_completa': fix_text(texto_rtf.get_text(separator=' ', strip=True)) if texto_rtf else '',
        })

    movimento_atual = None
    movimento_dia_seguinte = None

    for mov in movimentos:

        if mov.get('data') == data_planilha and mov.get('descricao_resumida') in MOVIMENTOS:
            movimento_atual = mov.get('descricao_resumida')

        if mov.get('data') == data_dia_seguinte and mov.get('descricao_resumida') in MOVIMENTOS:
            movimento_dia_seguinte = mov.get('descricao_resumida')

    movimento_atual = movimento_atual or 'Movimento não encontrado'
    movimento_dia_seguinte = movimento_dia_seguinte or 'Movimento não encontrado'

    return movimento_atual, movimento_dia_seguinte


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



def processar_processo(cookies, processo):
    numeros_expedientes, areas_judiciais = consulta_numero_expediente(cookies, processo)

    todos_movimentos_atuais = []
    todos_movimentos_dia_seguinte = []

    for numero_expediente in numeros_expedientes:
        movimento_atual, movimento_dia_seguinte = consulta_movimentos(cookies, numero_expediente, processo)
        todos_movimentos_atuais.append(movimento_atual)
        todos_movimentos_dia_seguinte.append(movimento_dia_seguinte)

    movimento_principal = 'Movimento não encontrado'
    movimento_principal_dia_seguinte = ''

    for movimento in todos_movimentos_atuais:
        if movimento != 'Movimento não encontrado':
            movimento_principal = movimento
            break

    for movimento in todos_movimentos_dia_seguinte:
        if movimento != 'Movimento não encontrado':
            movimento_principal_dia_seguinte = movimento
            break

    processo_atualizado = {
        **processo,
        'movimento': movimento_principal,
        'numero_expediente': numeros_expedientes[0] if numeros_expedientes else '',
        'area_judicial': areas_judiciais[0] if areas_judiciais else '',
        'movimento_dia_seguinte': movimento_principal_dia_seguinte
    }

    with lock:
        salvar_informacoes_no_json(processo_atualizado, NOME_ARQUIVO_PARA_SALVAR)


lock = threading.Lock()
processos = dados_planilha()
cookies = pegar_cookies()

with ThreadPoolExecutor(max_workers=10) as executor:
    future_to_processo = {
        executor.submit(processar_processo, cookies, processo): processo for processo in processos
    }

    for future in as_completed(future_to_processo):
        processo = future_to_processo[future]
        try:
            future.result()
        except Exception as e:
            print(f'Erro ao processar processo {processo}: {e}')

salvar_informacoes_no_excel()

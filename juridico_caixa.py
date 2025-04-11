import warnings

import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from ftfy import fix_text

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

  WebDriverWait(navegador, 50).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#textoBuscaSijur"))
  )

  cookies_lista = navegador.get_cookies()
  cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies_lista}
  cookies = '; '.join(f'{k}={v}' for k, v in cookies_dict.items())

  return cookies


def consulta_numero_expediente(cookies):
  url = "https://www.juridico.caixa.gov.br/?pg=busca"

  numero_do_processo = '00128073620194036315'

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

  # Encontrando a div que contém o número do processo
  div = soup.find('div', class_='tableCell', style=lambda x: x and 'white-space:nowrap' in x)

  return div.text.strip()


def consulta_movimentos(cookies, numero_expediente):
  url = f"https://www.juridico.caixa.gov.br/?pg=Expediente_movimentos&expediente={numero_expediente}&p_servidor=BDD7877WN001.corp.caixa.gov.br"

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

    movimentos.append({
      'codigo_movimento': fix_text(codigo.text.strip()) if codigo else '',
      'data': fix_text(data.text.strip()) if data else '',
      'descricao_resumida': fix_text(resumo.text.strip()) if resumo else '',
      'descricao_completa': fix_text(texto_rtf.get_text(separator=' ', strip=True)) if texto_rtf else ''
    })

  # Exemplo: imprimindo os três primeiros
  for mov in movimentos[:3]:
    print(mov)

cookies = pegar_cookies()
numero_expediente = consulta_numero_expediente(cookies)
movimentos = consulta_movimentos(cookies, numero_expediente)
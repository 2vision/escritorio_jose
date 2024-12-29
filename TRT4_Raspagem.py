import base64
import time
import warnings

import pandas as pd
import subprocess
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import StaleElementReferenceException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from anticaptchaofficial.imagecaptcha import *

warnings.simplefilter(action='ignore', category=FutureWarning)

API_KEY = "a89345c962e2eba448e571a6d0143363"


planilha_dados = pd.read_excel("TRT.xlsx", sheet_name="Plan1")

def para_planilha():
    # Criar um objeto writer especificando o arquivo Excel
    writer = pd.ExcelWriter("TRT.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace')

    # Atualizar os dados na aba "Ativos"
    planilha_dados.to_excel(writer, "Plan1", index=False)

    # Salvar as alterações
    writer.close()


edge_options = webdriver.EdgeOptions()
edge_options.set_capability("ms:edgeChromium", True)
edge_options.add_experimental_option("useAutomationExtension", False)
edge_options.add_argument("--disable-blink-features=AutomationControlled")
edge_options.add_argument("--disable-sync")
edge_options.add_argument("--disable-features=msEdgeEnableNurturingFramework")
edge_options.add_argument("--disable-popup-blocking")
edge_options.add_argument("--disable-infobars")
edge_options.add_argument("--disable-extensions")
edge_options.add_argument("--remote-allow-origins=*")
edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])

def obter_caminho_chromedriver():
    # Chame o gerenciador.py para obter o caminho do ChromeDriver
    result = subprocess.run(["python", "gerenciador.py"], capture_output=True, text=True)
    chromedriver_path = result.stdout.strip()
    return chromedriver_path


def resolver_captcha(captcha_base64):
    captcha_bytes = base64.b64decode(captcha_base64)

    # Enviar o captcha para o serviço
    response = requests.post(
        'https://api.anti-captcha.com/createTask',
        json={
            "clientKey": API_KEY,
            "task": {
                "type": "ImageToTextTask",
                "body": captcha_base64,
                "phrase": False,
                "case": False,
                "numeric": False,
                "math": 0,
                "minLength": 0,
                "maxLength": 0
            }
        }
    )

    task_id = response.json().get('taskId')
    if not task_id:
        raise Exception("Falha ao criar a tarefa de captcha")

    # Aguardar a solução do captcha
    while True:
        time.sleep(5)  # Aguardar 5 segundos entre as consultas
        result = requests.post(
            'https://api.anti-captcha.com/getTaskResult',
            json={"clientKey": API_KEY, "taskId": task_id}
        ).json()

        if result.get('status') == 'ready':
            return result['solution']['text']

prefs = {
    'user_experience_metrics': {
        'personalization_data_consent_enabled': True
    }
}

edge_options.add_experimental_option('prefs', prefs)

# Crie uma instância do Microsoft EdgeDriver com o serviço, opções e caminho do driver configurados
navegador = webdriver.Edge(options=edge_options)

navegador.maximize_window()

navegador.get("https://pje.trt4.jus.br/primeirograu/login.seam")


input("Pressione Enter após realizar o login...")

tres_pontinhos = navegador.find_element(By.XPATH, '//*[@id="botao-menu"]')
navegador.execute_script("arguments[0].click();", tres_pontinhos)

time.sleep(1)
consultar = navegador.find_element(By.XPATH, '//*[@id="menu-item-1"]/pje-menu-sobreposto/div[2]/pje-item-menu-sobreposto/div[3]/div/div/div[2]')
navegador.execute_script("arguments[0].click();", consultar)
time.sleep(1)
consultar_processos_de_terceiros = navegador.find_element(By.XPATH, '//*[@id="menu-item-1"]/pje-menu-sobreposto/div[2]/pje-item-menu-sobreposto/div[3]/div/a/div[2]')
navegador.execute_script("arguments[0].click();", consultar_processos_de_terceiros)


navegador.switch_to.window(navegador.window_handles[-1])

WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="nrProcessoInput"]')))

processos_nao_verificados = planilha_dados['CPF/CNPJ'].isnull().sum()

processos_verificados = 1

for indice_linha, linha in planilha_dados.iterrows():

    try:
        if pd.isna(linha["CPF/CNPJ"]):

            print(f"Processando verificados: {processos_verificados}/{processos_nao_verificados}")

            processos_verificados += 1

            num_processo = linha["Número do Processo"][:25]

            navegador.find_element(By.XPATH,'//*[@id="nrProcessoInput"]').send_keys(num_processo)
            consulta_processo = navegador.find_element(By.XPATH, '//*[@id="btnPesquisar"]')
            navegador.execute_script("arguments[0].click();", consulta_processo)

            WebDriverWait(navegador, 25).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="painelForm"]/form')))

            captcha_element = navegador.find_element(By.ID, "imagemCaptcha")

            # Obtém o valor do atributo 'src' da imagem
            captcha_base64 = captcha_element.get_attribute("src")

            # Remove a parte 'data:image/png;base64,'
            captcha_base64 = captcha_base64.split(',')[1]

            # Decodifica a string base64 para bytes
            captcha_bytes = base64.b64decode(captcha_base64)

            captcha_text = resolver_captcha(captcha_base64)

            # Preencher o campo de captcha com a solução
            navegador.find_element(By.XPATH, '//*[@id="captchaInput"]').send_keys(captcha_text)

            # Rolar até o botão e clicar para enviar o captcha
            btn_enviar = navegador.find_element(By.ID, "btnEnviar")
            navegador.execute_script("arguments[0].scrollIntoView(true);", btn_enviar)

            navegador.execute_script("arguments[0].click();", btn_enviar)

            WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="painel-titulo"]')))

            wait = WebDriverWait(navegador, 10)

            h1_element = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="painel-titulo"]//*[@id="titulo-detalhes"]//h1')))

            # Rolando até o elemento caso esteja fora da visão
            actions = ActionChains(navegador)
            actions.move_to_element(h1_element).perform()

            # Tentando clicar no elemento
            h1_element.click()

            # Capturando os dados com XPath
            orgao_julgador = navegador.find_element(By.XPATH,"//dt[text()='Órgão julgador:']/following-sibling::dd").text
            data_distribuicao = navegador.find_element(By.XPATH,"//dt[text()='Distribuído:']/following-sibling::dd").text
            valor_causa = navegador.find_element(By.XPATH,"//dt[text()='Valor da causa:']/following-sibling::dd").text
            valor_causa = valor_causa.replace('R$', '').replace('.', '').replace(',', '.').strip()
            valor_causa = float(valor_causa)
            planilha_dados.at[indice_linha, 'Valor da Causa'] = valor_causa

            assuntos = navegador.find_elements(By.XPATH,
                                               "//dt[text()='Assunto(s):']/following-sibling::dd[@class='ng-star-inserted']")

            assuntos_list = [assunto.text.strip() for assunto in assuntos]
            assuntos_separados = ', '.join(assuntos_list)

            polo_passivo = navegador.find_element(By.XPATH,
                                                  "//div[@class='coluna-polo']//h3[contains(text(), 'Polo passivo')]")

            # Verifica se o título do polo é "Polo Passivo"
            if polo_passivo:
                # Captura a seção de reclamados diretamente na coluna do polo passivo
                reclamados_section = polo_passivo.find_elements(By.XPATH, "following::pje-parte-processo//ul")

                nomes_reclamados = []
                documentos_reclamados = []

                # Itera sobre cada seção de reclamados
                for ul in reclamados_section:
                    # Localiza todos os reclamados dentro da seção
                    for reclamado in ul.find_elements(By.XPATH, ".//li[contains(@class, 'partes-corpo')]"):
                        # Capturando o nome do reclamado
                        nome = reclamado.find_element(By.XPATH,
                                                      ".//span[@class='nome-parte parte-documento-valido']").text.strip()
                        nomes_reclamados.append(nome)

                        # Capturando o CNPJ ou CPF do reclamado
                        documento_element = reclamado.find_element(By.XPATH,
                                                                   ".//span[contains(text(), 'CPJ:') or contains(text(), 'CPF:')]")
                        documento = documento_element.text.replace('CPJ: ', '').replace('CPF: ', '').strip()
                        documentos_reclamados.append(documento)

            if len(nomes_reclamados) > 0:

                planilha_dados.at[indice_linha, 'Reclamado'] = nomes_reclamados[0]
                planilha_dados.at[indice_linha, 'CPF/CNPJ'] = documentos_reclamados[0]
                planilha_dados.at[indice_linha, 'Orgão Julgador'] = orgao_julgador
                planilha_dados.at[indice_linha, 'Data de distribuição'] = data_distribuicao
                planilha_dados.at[indice_linha, 'Valor da Causa'] = valor_causa
                planilha_dados.at[indice_linha, 'Assuntos'] = assuntos_separados

                # Se houver mais de um reclamado, adiciona cada um em uma nova linha
                for i in range(1, len(nomes_reclamados)):  # Começa do segundo reclamado
                    # Criando uma nova linha para cada reclamado
                    new_row = pd.DataFrame({
                        'Reclamado': [nomes_reclamados[i]],
                        'Número do Processo': [num_processo],
                        'Reclamante': [linha["Reclamante"]],
                        'Orgão Julgador': [orgao_julgador],
                        'Data de distribuição': [data_distribuicao],
                        'Valor da Causa': [valor_causa],
                        'Assuntos': [assuntos_separados],  # Corrigido para 'assuntos_separados'
                        'CPF/CNPJ': [documentos_reclamados[i]]
                    })
                    # Inserindo a nova linha no DataFrame
                    planilha_dados = pd.concat([planilha_dados, new_row], ignore_index=True)

                para_planilha()

            navegador.find_element(By.XPATH, '//*[@id="btnFecharDadosProcessos"]/i').click()
            time.sleep(0.5)
            voltar = navegador.find_element(By.XPATH, "/html/body/pje-root/div[2]/pje-menu-lateral/div/ul/li[1]/a")
            navegador.execute_script("arguments[0].click();", voltar)



    except:
        time.sleep(0.5)
        planilha_dados.at[indice_linha, 'CPF/CNPJ'] = 'Segredo de justiça, Processo não pode ser aberto ou Falha no Captcha'
        voltar = navegador.find_element(By.XPATH, "/html/body/pje-root/div[2]/pje-menu-lateral/div/ul/li[1]/a")
        navegador.execute_script("arguments[0].click();", voltar)

        para_planilha()
        continue


navegador.quit()
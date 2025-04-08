import datetime
import math
import pandas as pd
import time
import shutil
import subprocess

from bs4 import BeautifulSoup
import re
from tkinter import *
from selenium import webdriver
from selenium.common import StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


planilha_dados = pd.read_excel("Excel/Consulta_TJPB.xlsx", sheet_name="Plan1")


def para_planilha():
    planilha_dados.to_excel("Excel/Consulta_TJPB.xlsx", sheet_name="Plan1", index=False)

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument('log-level=3')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")

# Obtenha o caminho do ChromeDriver do gerenciador.py
def obter_caminho_chromedriver():
    # Chame o gerenciador.py para obter o caminho do ChromeDriver
    result = subprocess.run(["python", "gerenciador.py"], capture_output=True, text=True)
    chromedriver_path = result.stdout.strip()
    return chromedriver_path

# Crie uma instância do ChromeDriver configurando o caminho diretamente
chrome_options.binary_location = obter_caminho_chromedriver()
navegador = webdriver.Chrome(options=chrome_options)

navegador.maximize_window()

termos_de_pesquisa = ["Execução Fiscal"]

navegador.get("https://pje.tjpb.jus.br/pje/login.seam?loginComCertificado=false")

input("Pressione Enter após realizar o login...")

data_fixa_inicial = "17/03/2025"
data_fixa_final = "17/03/2025"
valor_acao_fixa = "40000"

bancos_para_verificar = [
    "UNIAO FEDERAL - FAZENDA NACIONAL",
    "CONSELHO REGIONAL DE EDUCACAO FISICA DA 4 REGIAO",
    "INSTITUTO BRASILEIRO DO MEIO AMBIENTE E DOS RECURSOS NATURAIS RENOVAVEIS - IBAMA",
    "UNIAO",
    "CONSELHO",
    "FEDERAL",
    "FAZENDA",
    "MUNICIPIO",
    "ESTADO"
]
WebDriverWait(navegador, 300).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='barraSuperiorPrincipal']/div/div[1]/ul/li/a")))


verificado = planilha_dados["Nº do Processo"].any()

if not verificado:

    try:
        navegador.find_element(By.XPATH, '//*[@id="j_id127"]/span/i').click()

    except:
        pass

    navegador.find_element(By.XPATH, '//*[@id="barraSuperiorPrincipal"]/div/div[1]/ul/li/a').click()
    time.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="menu"]/div[2]/ul/li[1]').click()
    time.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="menu"]/div[2]/ul/li[1]/div/ul/li[4]').click()
    time.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="menu"]/div[2]/ul/li[1]/div/ul/li[4]/div/ul/li').click()

    for classe in termos_de_pesquisa:
        try:
            time.sleep(4)

            navegador.refresh()

            print(f'Iniciando busca de {classe} ...')

            time.sleep(2)
            classe_judicial_inserir = navegador.find_element(By.XPATH, '//input[@alt="Classe judicial"]').send_keys(
                classe)

            data_input_element = navegador.find_element(By.XPATH,
                                                        '//*[@id="fPP:dataAutuacaoDecoration:dataAutuacaoInicioInputDate"]')
            data_final_input_element = navegador.find_element(By.XPATH,
                                                              '//*[@id="fPP:dataAutuacaoDecoration:dataAutuacaoFimInputDate"]')
            valor_causa_input_element = navegador.find_element(By.XPATH,
                                                               '//*[@id="fPP:valorDaCausaDecoration:valorCausaInicial"]')

            script = f"arguments[0].value = '{data_fixa_inicial}';"
            navegador.execute_script(script, data_input_element)

            script = f"arguments[0].value = '{data_fixa_final}';"
            navegador.execute_script(script, data_final_input_element)

            script = f"arguments[0].value = '{valor_acao_fixa}';"
            navegador.execute_script(script, valor_causa_input_element)

            botao_pesquisar_element = navegador.find_element(By.ID, 'fPP:searchProcessos')

            navegador.execute_script("arguments[0].click();", botao_pesquisar_element)

            WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.CLASS_NAME, "rich-table-cell")))

            pagina_atual = 1
            total_paginas = navegador.find_element(By.XPATH, '//*[@id="fPP:processosTable:j_id464"]/div[2]/span').text
            total_paginas = total_paginas.split(" ")[0]
            total_paginas = math.ceil(int(total_paginas)/20)


            if total_paginas >= 10:

                num_pages = 15
            else:
                num_pages = total_paginas + 5


            while True:

                wait = WebDriverWait(navegador, 15)
                wait.until(EC.invisibility_of_element_located(
                    (By.XPATH, '//*[@id="_viewRoot:status.start"]/div/div[2]/div/div')))

                tabela_bancos = navegador.find_element(By.XPATH, "//*[@id='fPP:processosTable:tb']")
                colunas = tabela_bancos.find_elements(By.XPATH, ".//tr")

                for processos in colunas:
                    td_elements2 = processos.find_elements(By.TAG_NAME, "td")
                    numero_processo_banco = td_elements2[1].text
                    data_processo_banco = td_elements2[4].text
                    classe_judicial_banco = td_elements2[5].text
                    polo_ativo = td_elements2[6].text
                    povo_passivo = td_elements2[7].text

                    prox_linha = planilha_dados["Nº do Processo"].last_valid_index()

                    if pd.isna(prox_linha):
                        prox_linha = 0

                    else:
                        prox_linha += 1

                    if (
                            not any(banco in povo_passivo for banco in
                                    bancos_para_verificar) and
                            numero_processo_banco != '' and
                            data_processo_banco != '' and
                            classe_judicial_banco != '' and
                            polo_ativo != '' and
                            povo_passivo != ''
                    ):
                        planilha_dados.loc[prox_linha, "Nº do Processo"] = numero_processo_banco
                        planilha_dados.loc[prox_linha, "Data da distribuição"] = data_processo_banco
                        planilha_dados.loc[prox_linha, "Classe Judicial"] = classe_judicial_banco
                        planilha_dados.loc[prox_linha, "Polo Ativo"] = polo_ativo
                        planilha_dados.loc[prox_linha, "Cliente"] = povo_passivo
                        planilha_dados.loc[prox_linha, "Tribunal"] = "TJPB"

                        para_planilha()

                navegador.execute_script("window.scrollTo(0, -document.body.scrollHeight);")

                wait = WebDriverWait(navegador, 10)
                paginacao_element = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//*[@id='fPP:processosTable:scTabela_table']/tbody/tr/td[" + str(num_pages) + "]")))

                # Executa um script JavaScript para clicar no elemento de paginação
                script = "arguments[0].click();"
                navegador.execute_script(script, paginacao_element)

                if pagina_atual == total_paginas:
                    break
                pagina_atual += 1
        except:
            continue

print(f'Busca por processos finalizada!')

total_de_processos = len(planilha_dados[
                             (planilha_dados["Nº do Processo"].notna())])

total_processed = 0

navegador.refresh()

print(f'Começando agora análise de dados, processo para verificação {total_de_processos}')

try:
    navegador.find_element(By.XPATH, '//*[@id="j_id179"]/input[1]').click()

except:
    pass

try:
    navegador.find_element(By.XPATH, '//*[@id="j_id127"]/span/i').click()

except:
    pass

navegador.find_element(By.XPATH, '//*[@id="barraSuperiorPrincipal"]/div/div[1]/ul/li/a').click()
time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="menu"]/div[2]/ul/li[1]').click()
time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="menu"]/div[2]/ul/li[1]/div/ul/li[4]').click()
time.sleep(2)
navegador.find_element(By.XPATH, '//*[@id="menu"]/div[2]/ul/li[1]/div/ul/li[4]/div/ul/li').click()

for indice_linha2, linha2 in planilha_dados.iterrows():

    try:
        if pd.isna(linha2["CPF/CNPJ"]):

            total_processed += 1
            print(f"Processos verificados {total_processed}/{total_de_processos}")

            num_processo = linha2["Nº do Processo"]

            num_processo = num_processo.split("-")
            num_01 = num_processo[0]
            num_processo = num_processo[1].split(".")
            num_02 = num_processo[0]
            num_03 = num_processo[1]
            num_04 = num_processo[2]
            num_05 = num_processo[3]
            num_06 = num_processo[4]

            navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            input_element = navegador.find_element(By.ID, 'fPP:numeroProcesso:numeroSequencial')
            input_element.clear()

            input_element2 = navegador.find_element(By.XPATH, '//*[@id="fPP:numeroProcesso:numeroDigitoVerificador"]')
            input_element2.clear()

            input_element3 = navegador.find_element(By.XPATH, '//*[@id="fPP:numeroProcesso:Ano"]')
            input_element3.clear()

            input_element6 = navegador.find_element(By.XPATH, '//*[@id="fPP:numeroProcesso:NumeroOrgaoJustica"]')
            input_element6.clear()


            input_element = navegador.find_element(By.ID, 'fPP:numeroProcesso:numeroSequencial')
            input_element.send_keys(num_01)

            input_element2 = navegador.find_element(By.XPATH, '//*[@id="fPP:numeroProcesso:numeroDigitoVerificador"]')
            input_element2.send_keys(num_02)

            input_element3 = navegador.find_element(By.XPATH, '//*[@id="fPP:numeroProcesso:Ano"]')
            input_element3.send_keys(num_03)

            input_element6 = navegador.find_element(By.XPATH, '//*[@id="fPP:numeroProcesso:NumeroOrgaoJustica"]')
            input_element6.send_keys(num_06)

            navegador.execute_script("window.scrollTo(0, -document.body.scrollHeight);")

            botao_pesquisar_element = navegador.find_element(By.ID, 'fPP:searchProcessos')

            navegador.execute_script("arguments[0].click();", botao_pesquisar_element)

            time.sleep(3)

            wait = WebDriverWait(navegador, 10)
            element = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'btn-link.btn-condensed')))

            navegador.execute_script("arguments[0].click();", element)

            wait = WebDriverWait(navegador, 10)

            alert = wait.until(EC.alert_is_present())

            alert.accept()

            time.sleep(3)

            navegador.switch_to.window(navegador.window_handles[-1])

            wait.until(EC.presence_of_all_elements_located((By.XPATH, '//body[not(@loading) or @loading="false"]')))

            html = navegador.page_source

            soup = BeautifulSoup(html, 'html.parser')

            try:
                element_dt = soup.find('dt', string=re.compile(r'Valor da causa', re.IGNORECASE))

                if element_dt:
                    value_element = element_dt.find_next('dd')
                    valor_da_causa = value_element.get_text(strip=True)
                    valor_da_causa = valor_da_causa.replace("R$", "").replace(".", "")
                    planilha_dados.loc[indice_linha2, "Valor Causa"] = valor_da_causa
            except:

                planilha_dados.loc[indice_linha2, "Valor da Causa"] = "Valor da causa não encontrado"

            try:
                # Encontre o elemento <div> com id="poloPassivo"
                polo_passivo_div = soup.find('div', id='poloPassivo')

                # Se o elemento <div> for encontrado, encontre o <span> com a informação de CPF ou CNPJ
                if polo_passivo_div:
                    cpf_cnpj_span = polo_passivo_div.find('span', string=re.compile(r'(CPF|CNPJ):'))

                    if cpf_cnpj_span:
                        cpf_cnpj_text = cpf_cnpj_span.get_text(strip=True)
                        cpf_cnpj = re.search(r'(CPF|CNPJ): (\S+)', cpf_cnpj_text).group(2)
                        planilha_dados.loc[indice_linha2, "CPF/CNPJ"] = cpf_cnpj

            except:
                planilha_dados.loc[indice_linha2, "CPF/CNPJ"] = "CPF não encontrado"

            try:
                element_dt_assunto = soup.find('dt', string=re.compile(r'Assunto', re.IGNORECASE))

                if element_dt_assunto:
                    assunto_element = element_dt_assunto.find_next('dd')
                    assunto = assunto_element.get_text(strip=True).replace("<br>", " / ")
                    planilha_dados.loc[indice_linha2, "Assunto"] = assunto
            except:
                planilha_dados.loc[indice_linha2, "Assunto"] = "Assunto não encontrado"

            try:
                element_dt_orgao = soup.find('dt', string=re.compile(r'Órgão Julgador', re.IGNORECASE))

                # Se o elemento <dt> for encontrado, pegue o próximo elemento <dd> que contém o órgão julgador
                if element_dt_orgao:
                    orgao_element = element_dt_orgao.find_next('dd')
                    orgao_julgador = orgao_element.get_text(strip=True)
                    planilha_dados.loc[indice_linha2, "Órgão Julgador"] = orgao_julgador
            except:
                planilha_dados.loc[indice_linha2, "Órgão Julgador"] = "Órgão Julgador não encontrado"


            try:
                tem_advogado = navegador.find_element(By.XPATH, "//*[@id='poloPassivo']/table/tbody/tr/td/ul/li/small/span/span")
                planilha_dados.loc[indice_linha2, "Advogado"] = "Já tem advogado"
            except:
                planilha_dados.loc[indice_linha2, "Advogado"] = "Não tem advogado"

            para_planilha()

            navegador.execute_script("window.close();")

            navegador.switch_to.window(navegador.window_handles[0])

    except StaleElementReferenceException:

        if len(navegador.window_handles) == 2:

            time.sleep(2)
            navegador.switch_to.window(navegador.window_handles[1])

            time.sleep(2)

            navegador.close()

            time.sleep(2)
            navegador.switch_to.window(navegador.window_handles[0])

        planilha_dados.loc[indice_linha2, "CPF/CNPJ"] = "erro na linha"

        navegador.refresh()

        time.sleep(5)

        para_planilha()

        continue

planilha_dados = planilha_dados[(planilha_dados["CPF/CNPJ"].notna()) & (planilha_dados["CPF/CNPJ"] != "CPF não encontrado")]

# Salve a planilha atualizada no Excel
para_planilha()

print(f'Pesquisa finalizada até a próxima !')
navegador.close()
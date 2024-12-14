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
from webdriver_manager.chrome import ChromeDriverManager

planilha_dados = pd.read_excel("Excel/TRF3.xlsx", sheet_name="Plan1")

def para_planilha():
    planilha_dados.to_excel("Excel/TRF3.xlsx", sheet_name="Plan1", index=False)

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument('log-level=3')  # para ignorar warnings
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # para ignorar warnings

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

navegador.get("https://pje1g.trf3.jus.br/pje/login.seam?loginComCertificado=false")


data_fixa_inicial = "20/03/2023"
data_fixa_final = "20/03/2023"
valor_acao_fixa = "10.000,00"
valor_acao_final = "10.000,00"


logo = r"""

_  _ _ ____ _ ____ _  _    ___ ____ ____ _  _ _  _ ____ _    ____ ____ _   _ 
|  | | [__  | |  | |\ |     |  |___ |    |__| |\ | |  | |    |  | | __  \_/  
 \/  | ___] | |__| | \|     |  |___ |___ |  | | \| |__| |___ |__| |__]   |   

    """
print(logo)

bancos_para_verificar = [
    "UNIAO FEDERAL - FAZENDA NACIONAL",
    "CONSELHO REGIONAL DE EDUCACAO FISICA DA 4 REGIAO",
    "INSTITUTO BRASILEIRO DO MEIO AMBIENTE E DOS RECURSOS NATURAIS RENOVAVEIS - IBAMA",
    "UNIAO",
    "CONSELHO",
    "FEDERAL",
    "FAZENDA"
]

WebDriverWait(navegador, 300).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='barraSuperiorPrincipal']/div/div[1]/ul/li/a")))


verificado = planilha_dados["Nº do Processo"].any()

if not verificado:

    def formatar_data(entry):
        current_text = entry.get()
        length = len(current_text)

        if length == 2 or length == 5:
            entry.insert(len(current_text), '/')


    def atualizar_data_inicial():
        global data_fixa_inicial
        data_fixa_inicial = data_entry_inicial.get()
        janela_data_inicial.destroy()


    janela_data_inicial = Tk()
    janela_data_inicial.title("Digite a data de Início")

    largura_janela = 350
    altura_janela = 100

    largura_tela = janela_data_inicial.winfo_screenwidth()
    altura_tela = janela_data_inicial.winfo_screenheight()

    posicao_x = int(largura_tela / 2 - largura_janela / 2)
    posicao_y = int(altura_tela / 2 - altura_janela / 2)

    janela_data_inicial.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

    data_label_inicial = Label(janela_data_inicial, text="Digite a data de início da pesquisa (DD/MM/AAAA):")
    data_label_inicial.pack()

    data_entry_inicial = Entry(janela_data_inicial)
    data_entry_inicial.pack()

    # Configurando para formatar a data enquanto o usuário digita
    data_entry_inicial.bind("<KeyRelease>", lambda event: formatar_data(data_entry_inicial))

    ok_button_inicial = Button(janela_data_inicial, text="OK", command=atualizar_data_inicial)
    ok_button_inicial.pack()

    janela_data_inicial.mainloop()


    def atualizar_data_final():
        global data_fixa_final
        data_fixa_final = data_entry_final.get()
        janela_data_final.destroy()


    janela_data_final = Tk()
    janela_data_final.title("Digite a data Final")

    largura_janela = 350
    altura_janela = 100

    largura_tela = janela_data_final.winfo_screenwidth()
    altura_tela = janela_data_final.winfo_screenheight()

    posicao_x = int(largura_tela / 2 - largura_janela / 2)
    posicao_y = int(altura_tela / 2 - altura_janela / 2)

    janela_data_final.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

    data_label_final = Label(janela_data_final, text="Digite a data final da pesquisa (DD/MM/AAAA):")
    data_label_final.pack()

    data_entry_final = Entry(janela_data_final)
    data_entry_final.pack()

    # Configurando para formatar a data enquanto o usuário digita
    data_entry_final.bind("<KeyRelease>", lambda event: formatar_data(data_entry_final))

    ok_button_final = Button(janela_data_final, text="OK", command=atualizar_data_final)
    ok_button_final.pack()

    janela_data_final.mainloop()

    def atualiza_valor():
        global valor_acao_fixa
        valor_acao_fixa = valor_entry.get()
        janela_valor.destroy()


    janela_valor = Tk()
    janela_valor.title("Valor mínimo da ação")

    largura_janela = 350
    altura_janela = 100

    largura_tela = janela_valor.winfo_screenwidth()
    altura_tela = janela_valor.winfo_screenheight()

    posicao_x = int(largura_tela / 2 - largura_janela / 2)
    posicao_y = int(altura_tela / 2 - altura_janela / 2)

    janela_valor.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

    valor_label = Label(janela_valor, text="Digite o valor mínimo da ação")
    valor_label.pack()

    valor_entry = Entry(janela_valor)
    valor_entry.pack()

    valor_button = Button(janela_valor, text="OK", command=atualiza_valor)
    valor_button.pack()

    janela_valor.mainloop()

    def atualiza_valor_final():
        global valor_acao_fixa_final
        valor_acao_fixa_final = valor_entry_final.get()
        janela_valor_final.destroy()


    janela_valor_final = Tk()
    janela_valor_final.title("Valor da máximo da ação")

    largura_janela = 350
    altura_janela = 100

    largura_tela = janela_valor_final.winfo_screenwidth()
    altura_tela = janela_valor_final.winfo_screenheight()

    posicao_x = int(largura_tela / 2 - largura_janela / 2)
    posicao_y = int(altura_tela / 2 - altura_janela / 2)

    janela_valor_final.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

    valor_label_final = Label(janela_valor_final, text="Digite o valor máximo da ação")
    valor_label_final.pack()

    valor_entry_final = Entry(janela_valor_final)
    valor_entry_final.pack()

    valor_button_final= Button(janela_valor_final, text="OK", command=atualiza_valor_final)
    valor_button_final.pack()

    janela_valor_final.mainloop()

    # CPF 823.111.290-15
    # SENHA: Vmb605606!

    try:
        navegador.find_element(By.XPATH, '//*[@id="j_id127"]/span/i').click()

    except:
        pass

    navegador.find_element(By.XPATH, "//*[@id='barraSuperiorPrincipal']/div/div[1]/ul/li/a").click()
    time.sleep(2)
    navegador.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/a").click()
    time.sleep(2)
    navegador.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]/a").click()
    time.sleep(2)
    navegador.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]/div/ul/li/a").click()

    for classe in termos_de_pesquisa:
        try:

            time.sleep(4)

            navegador.refresh()

            print(f'Iniciando busca de {classe} ...')

            time.sleep(4)
            classe_judicial_inserir = navegador.find_element(By.XPATH, '//input[@alt="Classe judicial"]').send_keys(
                classe)

            data_input_element = navegador.find_element(By.XPATH,
                                                        '//*[@id="fPP:dataAutuacaoDecoration:dataAutuacaoInicioInputDate"]')
            data_final_input_element = navegador.find_element(By.XPATH,
                                                              '//*[@id="fPP:dataAutuacaoDecoration:dataAutuacaoFimInputDate"]')
            valor_causa_input_element = navegador.find_element(By.XPATH,
                                                               '//*[@id="fPP:valorDaCausaDecoration:valorCausaInicial"]')

            valor_causa_final_input_element = navegador.find_element(By.XPATH,
                                                               '//*[@id="fPP:valorDaCausaDecoration:valorCausaFinal"]')


            script = f"arguments[0].value = '{data_fixa_inicial}';"
            navegador.execute_script(script, data_input_element)

            script = f"arguments[0].value = '{data_fixa_final}';"
            navegador.execute_script(script, data_final_input_element)

            script = f"arguments[0].value = '{valor_acao_fixa}';"
            navegador.execute_script(script, valor_causa_input_element)

            script = f"arguments[0].value = '{valor_acao_fixa_final}';"
            navegador.execute_script(script, valor_causa_final_input_element)

            botao_pesquisar_element = navegador.find_element(By.ID, 'fPP:searchProcessos')

            # Execute o script para clicar no botão
            navegador.execute_script("arguments[0].click();", botao_pesquisar_element)

            # Execute o script para clicar no botão
            navegador.execute_script("arguments[0].click();", botao_pesquisar_element)

            WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.CLASS_NAME, "rich-table-cell")))

            def proxima_pagina(navegador):
                try:

                    time.sleep(3)
                    botao_proxima_pagina = navegador.find_element(By.XPATH,
                                                                  "//td[contains(@class, 'rich-datascr-button') and text()='»' and not(contains(@class, 'rich-datascr-button-dsbld'))]")
                    navegador.execute_script("arguments[0].scrollIntoView();", botao_proxima_pagina)
                    time.sleep(1)
                    botao_proxima_pagina.click()
                    return True
                except:
                    return False
            #
            #
            # pagina_atual = 1
            # total_paginas = navegador.find_element(By.XPATH, '//*[@id="fPP:processosTable:j_id411"]/div[2]/span').text
            #
            # if "Pesquisa limitada ao retorno de 100 processos" in total_paginas:
            #     total_paginas = math.ceil(100 / 20)
            #
            # else:
            #     total_paginas = total_paginas.split(" ")[0]
            #     total_paginas = math.ceil(int(total_paginas)/20)
            #
            #
            # if total_paginas >= 10:
            #
            #     num_pages = 15
            # else:
            #     num_pages = total_paginas + 5


            while True:

                time.sleep(5)


                tabela_bancos = navegador.find_element(By.XPATH, "//*[@id='fPP:processosTable:tb']")
                colunas = tabela_bancos.find_elements(By.XPATH, ".//tr")



                for processos in colunas:
                    time.sleep(1)
                    td_elements2 = processos.find_elements(By.TAG_NAME, "td")
                    numero_processo_banco = td_elements2[1].get_attribute("innerText")
                    data_processo_banco = td_elements2[4].get_attribute("innerText")
                    classe_judicial_banco = td_elements2[5].get_attribute("innerText")
                    polo_ativo = td_elements2[6].get_attribute("innerText")
                    povo_passivo = td_elements2[7].get_attribute("innerText")

                    prox_linha = planilha_dados["Nº do Processo"].last_valid_index()

                    if pd.isna(prox_linha):
                        prox_linha = 0

                    else:
                        prox_linha += 1

                    time.sleep(1)

                    planilha_dados.loc[prox_linha, "Nº do Processo"] = numero_processo_banco
                    planilha_dados.loc[prox_linha, "Data da distribuição"] = data_processo_banco
                    planilha_dados.loc[prox_linha, "Classe Judicial"] = classe_judicial_banco
                    planilha_dados.loc[prox_linha, "Polo Ativo"] = polo_ativo
                    planilha_dados.loc[prox_linha, "Cliente"] = povo_passivo
                    planilha_dados.loc[prox_linha, "Tribunal"] = "TRF3"

                    time.sleep(0.5)

                    para_planilha()

                time.sleep(1)
                if not proxima_pagina(navegador):
                    break

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

navegador.find_element(By.XPATH, "//*[@id='barraSuperiorPrincipal']/div/div[1]/ul/li/a").click()
time.sleep(2)
navegador.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/a").click()
time.sleep(2)
navegador.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]/a").click()
time.sleep(2)
navegador.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]/div/ul/li/a").click()

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

            botao_pesquisar_element = navegador.find_element(By.ID, 'fPP:searchProcessos')

            # Execute o script para clicar no botão
            navegador.execute_script("arguments[0].click();", botao_pesquisar_element)

            time.sleep(3)

            # Aguarde até que o elemento seja visível na página
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
                # Encontre o elemento <dt> com o texto 'Valor da causa'
                element_dt = soup.find('dt', string=re.compile(r'Valor da causa', re.IGNORECASE))

                # Se o elemento <dt> for encontrado, pegue o próximo elemento <dd> que contém o valor
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
                # Encontre o elemento <dt> com o texto 'Assunto'
                element_dt_assunto = soup.find('dt', string=re.compile(r'Assunto', re.IGNORECASE))

                # Se o elemento <dt> for encontrado, pegue o próximo elemento <dd> que contém o assunto
                if element_dt_assunto:
                    assunto_element = element_dt_assunto.find_next('dd')
                    assunto = assunto_element.get_text(strip=True).replace("<br>", " / ")
                    planilha_dados.loc[indice_linha2, "Assunto"] = assunto
            except:
                planilha_dados.loc[indice_linha2, "Assunto"] = "Assunto não encontrado"

            try:
                # Encontre o elemento <dt> com o texto 'Órgão Julgador'
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

            para_planilha()

            navegador.close()

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
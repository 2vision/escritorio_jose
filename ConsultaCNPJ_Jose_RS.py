import datetime
import locale
import subprocess
import warnings

import pandas as pd
import time
from tkinter import *
from selenium import webdriver
from selenium.common import NoSuchElementException, UnexpectedAlertPresentException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select


warnings.simplefilter(action='ignore', category=FutureWarning)

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')


planilha_dados = pd.read_excel("Excel/CNPJ_RS.xlsx", sheet_name="Plan1")

def para_planilha():
    planilha_dados.to_excel("Excel/CNPJ_RS.xlsx", sheet_name="Plan1", index=False)


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

prefs = {
    'user_experience_metrics': {
        'personalization_data_consent_enabled': True
    }
}

edge_options.add_experimental_option('prefs', prefs)

# Crie uma instância do Microsoft EdgeDriver com o serviço, opções e caminho do driver configurados
navegador = webdriver.Edge(options=edge_options)

navegador.get("https://eproc1g.tjrs.jus.br/eproc/externo_controlador.php")

navegador.maximize_window()

logo = r"""

_  _ _ ____ _ ____ _  _    ___ ____ ____ _  _ _  _ ____ _    ____ ____ _   _ 
|  | | [__  | |  | |\ |     |  |___ |    |__| |\ | |  | |    |  | | __  \_/  
 \/  | ___] | |__| | \|     |  |___ |___ |  | | \| |__| |___ |__| |__]   |   

    """
print(logo)

# navegador.find_element(By.ID, "txtUsuario").send_keys("rs091631")  # "rs091631"
# navegador.find_element(By.ID, "pwdSenha").send_keys("Sucesso2023!")  # "Sucesso2023!"
# navegador.find_element(By.ID, "sbmEntrar").click()

data_fixa = datetime.datetime.strptime("20/03/2023", '%d/%m/%Y')

verificado = (planilha_dados["CNPJ Verificado"] == "Feito").any()


try:

    elemento_tr0_espera = WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.ID, "tr0")))

    navegador.find_element(By.ID, "tr0").click()


except:

    pass

if not verificado:

    def atualizar_data():
        global data_fixa
        data_str = data_entry.get()
        data_fixa = datetime.datetime.strptime(data_str, '%d/%m/%Y')
        janela_data.destroy()


    janela_data = Tk()
    janela_data.title("Digite a data de inicio da pesquisa")

    largura_janela = 250
    altura_janela = 100

    largura_tela = janela_data.winfo_screenwidth()
    altura_tela = janela_data.winfo_screenheight()

    posicao_x = int(largura_tela / 2 - largura_janela / 2)
    posicao_y = int(altura_tela / 2 - altura_janela / 2)

    janela_data.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

    data_label = Label(janela_data, text="Digite a data de inicio da pesquisa (DD/MM/AAAA):")
    data_label.pack()

    data_entry = Entry(janela_data)
    data_entry.pack()

    ok_button = Button(janela_data, text="OK", command=atualizar_data)
    ok_button.pack()

    janela_data.mainloop()

    total_cnpj_consultado = 0

    total_de_cnpj = len(planilha_dados[
                            (planilha_dados["CNPJ"].notna())])

    WebDriverWait(navegador, 350).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="navbar"]/div/div[1]/div[4]/form/select')))

    print(f'Iniciando busca por CNPJ...')

    for indice_linha2, linha2 in planilha_dados.iterrows():
        try:

            total_cnpj_consultado += 1

            cnpj = linha2["CNPJ"]

            print(f'CNPJ Verificados {total_cnpj_consultado}/{total_de_cnpj}')

            entrar_primeiro_menu = WebDriverWait(navegador, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='main-menu']/li[5]/a"))
            )
            navegador.execute_script("arguments[0].click();", entrar_primeiro_menu)

            entrar_sub_menu = WebDriverWait(navegador, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-ul-3']/li/a"))
            )
            navegador.execute_script("arguments[0].click();", entrar_sub_menu)

            navegador.find_element(By.XPATH, "//*[@id='selTipoPesquisa']/option[3]").click()
            navegador.find_element(By.XPATH, "//*[@id='divStrDocParte']/dl/dd/input").clear()
            navegador.find_element(By.XPATH, "//*[@id='divStrDocParte']/dl/dd/input").send_keys(cnpj)

            entrar_na_consulta = WebDriverWait(navegador, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='sbmConsultar']"))
            )
            navegador.execute_script("arguments[0].click();", entrar_na_consulta)

            wait = WebDriverWait(navegador, 150)
            wait.until(EC.invisibility_of_element_located((By.ID, "divInfraAviso")))

            # select_element = navegador.find_element(By.ID, 'selInfraPaginacaoSuperior')
            #
            # # Criar objeto Select para interagir com o elemento select
            # select = Select(select_element)
            #
            # # Obter todas as opções do select
            # options = select.options
            #
            # # Determinar o último índice da lista de opções
            # ultimo_indice = len(options) - 1
            #
            # # Selecionar o último item da lista
            # select.select_by_index(ultimo_indice)
            #
            # wait = WebDriverWait(navegador, 150)
            # wait.until(EC.invisibility_of_element_located((By.ID, "divInfraAviso")))

            try:
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="divInfraExcecao"]/span')))

                planilha_dados.loc[indice_linha2, "CNPJ Verificado"] = "Feito"
                para_planilha()

                continue

            except:

                WebDriverWait(navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='divInfraAreaTabela']/tbody/tr")))

                tabela_bancos = navegador.find_elements(By.XPATH, "//*[@id='divInfraAreaTabela']/tbody/tr")

                contador = 0

                prox_linha = planilha_dados["Nº do Processo"].last_valid_index()

                if pd.isna(prox_linha):
                    prox_linha = 0

                else:
                    prox_linha += 1

                for processo_bancos in tabela_bancos[:200]:

                    contador += 1

                    td_elements2 = processo_bancos.find_elements(By.TAG_NAME, "td")

                    numero_processo_banco = td_elements2[0].text

                    data_processo_banco = td_elements2[1].text
                    classee_judicial = td_elements2[5].text
                    data_processo_banco = data_processo_banco.split()[0]
                    data_processo_banco_atualizada = datetime.datetime.strptime(data_processo_banco, '%d/%m/%Y')
                    nome_autor = td_elements2[3].text

                    if data_processo_banco_atualizada >= data_fixa:

                        planilha_dados.loc[prox_linha, "Nº do Processo"] = numero_processo_banco
                        planilha_dados.loc[prox_linha, "Data da distribuição"] = data_processo_banco
                        planilha_dados.loc[prox_linha, "Classe Judicial"] = classee_judicial

                        prox_linha += 1

                        para_planilha()

            planilha_dados.loc[indice_linha2, "CNPJ Verificado"] = "Feito"
            para_planilha()

        except:
            navegador.refresh()
            planilha_dados.loc[indice_linha2, "CNPJ Verificado"] = "Feito/Erro"
            continue


WebDriverWait(navegador, 350).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="navbar"]/div/div[1]/div[4]/form/select')))

print(f'Busca por processos finalizada!')

total_de_processos = len(planilha_dados[
                             (planilha_dados["Nº do Processo"].notna())])

total_processed = 0

print(f'Começando agora análise de dados, processo para verificação {total_de_processos}')

for indice_linha, linha in planilha_dados.iterrows():

    try:

        total_processed += 1
        print(f"Processos verificados {total_processed}/{total_de_processos}")

        if pd.isna(linha["CPF/CNPJ"]):
            num_processo = linha["Nº do Processo"]

            entrar_primeiro_menu = WebDriverWait(navegador, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='main-menu']/li[5]/a"))
            )
            navegador.execute_script("arguments[0].click();", entrar_primeiro_menu)

            entrar_sub_menu = WebDriverWait(navegador, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='menu-ul-3']/li/a"))
            )
            navegador.execute_script("arguments[0].click();", entrar_sub_menu)

            navegador.find_element(By.XPATH, "//*[@id='numNrProcesso']").clear()
            navegador.find_element(By.XPATH, "//*[@id='numNrProcesso']").send_keys(num_processo)
            entrar_na_consulta = WebDriverWait(navegador, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='sbmConsultar']"))

            )
            navegador.execute_script("arguments[0].click();", entrar_na_consulta)

            time.sleep(1)

            WebDriverWait(navegador, 15).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="divInfraBarraComandosSuperior"]')))

            if 'Segredo de Justiça (Nível 1)' in navegador.find_element(By.XPATH,'//*[@id="divInfraBarraComandosSuperior"]').text:

                planilha_dados.loc[indice_linha, "CPF/CNPJ"] = 'erro na linha'
                para_planilha()
                continue

            WebDriverWait(navegador, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="fldPartes"]')))

            try:
                # Obtendo o HTML da página atual
                html = navegador.page_source

                # Inicializando o Beautiful Soup
                soup = BeautifulSoup(html, 'html.parser')

                # Encontre a tag <td> que contém o texto "Valor da Causa"
                td_valor_causa = soup.find('td', string='Valor da Causa: ')

                # Se encontrarmos a tag <td>
                if td_valor_causa:
                    # Navegue para a tag <label> dentro dessa tag <td>
                    label_valor_causa = td_valor_causa.find_next('label', class_='infraLabelObrigatorio')

                    # Obtenha o valor da causa do texto do elemento <label>
                    valor_da_causa = label_valor_causa.get_text(strip=True).replace('Valor da Causa:', '').strip()

                    planilha_dados.loc[indice_linha, "Valor Causa"] = valor_da_causa
            except:

                planilha_dados.loc[indice_linha, "Valor Causa"] = "Não aparece valor da causa"

            try:
                nome_reu = navegador.find_element(By.ID, "spnNomeParteReu0").text
                planilha_dados.loc[indice_linha, "Cliente"] = nome_reu

            except:

                planilha_dados.loc[indice_linha, "Cliente"] = "Não Aparece o Nome"

            try:
                nome_autor = navegador.find_element(By.XPATH,
                                                    "//*[@id='tblPartesERepresentantes']/tbody/tr[2]/td[1]/span[1]").text
                planilha_dados.loc[indice_linha, "Polo Ativo"] = nome_autor

            except:

                planilha_dados.loc[indice_linha, "Polo Ativo"] = "Não Aparece o Nome"

            try:
                cpf_cliente = navegador.find_element(By.ID, "spnCpfParteReu0").text

                planilha_dados.loc[indice_linha, "CPF/CNPJ"] = cpf_cliente


            except NoSuchElementException:

                planilha_dados.loc[indice_linha, "CPF/CNPJ"] = "cliente sem cpf/cnpj"
                para_planilha()

                continue

            try:
                orgao_julgador = navegador.find_element(By.XPATH,
                                                         '//*[@id="txtOrgaoJulgador"]')
                planilha_dados.loc[indice_linha, "Órgão Julgador"] = orgao_julgador

            except:
                planilha_dados.loc[indice_linha, "Órgão Julgador"] = "Não encontrado"

            tr_element = soup.find('tr', {'data-assunto-principal': 'true'})

            # Verifique se o elemento <tr> foi encontrado
            if tr_element:
                # Encontre todos os elementos <td> dentro do <tr>
                td_elements = tr_element.find_all('td')

                # Certifique-se de que há pelo menos 2 elementos <td>
                if len(td_elements) >= 2:
                    # Pegue o conteúdo do segundo <td> (índice 1) dentro do <tr>
                    assunto = td_elements[1].get_text(strip=True)

                    planilha_dados.loc[indice_linha, "Assunto"] = assunto
                else:
                    planilha_dados.loc[indice_linha, "Assunto"] = "Não aparece o assunto"

            try:
                advogado_valido = navegador.find_element(By.XPATH,
                                                         "//*[@id='tblPartesERepresentantes']/tbody/tr[2]/td[2]/a")
                planilha_dados.loc[indice_linha, "Advogado"] = "Tem Advogado"

            except:
                planilha_dados.loc[indice_linha, "Advogado"] = "Sem Advogado"

            para_planilha()

    except UnexpectedAlertPresentException:

        planilha_dados.loc[indice_linha, "CPF/CNPJ"] = "erro na linha"
        navegador.refresh()
        para_planilha()
        continue

planilha_dados = planilha_dados[
    (planilha_dados["CPF/CNPJ"].notna()) &  # Não é NaN
    (planilha_dados["CPF/CNPJ"] != "CPF não encontrado") &  # Não contém essa mensagem
    (planilha_dados["CPF/CNPJ"].str.strip() != "erro na linha") &  # Não é vazio após remover espaços
    (planilha_dados["CPF/CNPJ"].str.lower() != "cliente sem cpf/cnpj")  # Não é "cliente sem cpf/cnpj"
]

para_planilha()

print(f'Pesquisa finalizada, dados enviados para a planilha CNPJ_RS!')
navegador.close()
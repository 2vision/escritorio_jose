import json
import locale
import os
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext

import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# === CONFIGURAÇÕES ===
NOME_ARQUIVO_PARA_SALVAR = 'Consulta JUSBR'


def api_jusbr(bearer_code, cnpj, paginacao=None):
    url_base = 'https://portaldeservicos.pdpj.jus.br/api/v2/processos'
    headers = {'Authorization': f'{bearer_code}'}
    query_params = {'cpfCnpjParte': cnpj}
    if paginacao:
        query_params['searchAfter'] = paginacao
    response = requests.get(url_base, headers=headers, params=query_params)
    return response.json() if response.status_code == 200 else None


def formatar_documento(documento):
    if documento:
        if len(documento) == 11:
            return f"{documento[:3]}.{documento[3:6]}.{documento[6:9]}-{documento[9:]}"
        elif len(documento) == 14:
            return f"{documento[:2]}.{documento[2:5]}.{documento[5:8]}/{documento[8:12]}-{documento[12:]}"
    return None


def salvar_informacoes_no_json(informacoes):
    dados = []
    if os.path.exists(f'{NOME_ARQUIVO_PARA_SALVAR}.json'):
        with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
            dados = json.load(f)
    dados.append(informacoes)
    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


def salvar_informacoes_no_excel():
    with open(f'{NOME_ARQUIVO_PARA_SALVAR}.json', 'r', encoding='utf-8') as f:
        dados = json.load(f)
    df = pd.DataFrame(dados)
    df.to_excel(f'{NOME_ARQUIVO_PARA_SALVAR}.xlsx', index=False, engine='openpyxl')
    os.remove(f'{NOME_ARQUIVO_PARA_SALVAR}.json')


def capturar_informacoes(processo, data_inicial):
    partes = processo.get('tramitacoes', [{}])[0].get('partes', [{}])
    classes = processo.get('tramitacoes', [{}])[0].get('classe', [{}])
    assunto = processo.get('tramitacoes', [{}])[0].get('assunto', [{}])[0].get("descricao")
    distribuicao = processo['tramitacoes'][0].get('dataHoraUltimaDistribuicao')
    data_distribuicao = datetime.strptime(distribuicao.split('.')[0], '%Y-%m-%dT%H:%M:%S') if distribuicao else None
    valor_acao = processo['tramitacoes'][0].get('valorAcao')
    estado_tribunal = processo.get('siglaTribunal')
    numero_processo = processo.get('numeroProcesso')
    nome_ativo, doc_ativo, nome_passivo, doc_passivo = None, None, None, None

    for parte in partes:
        if parte.get('polo') == 'ATIVO':
            nome_ativo = parte.get('nome')
            doc_ativo = parte.get('documentosPrincipais', [{}])[0].get('numero')
        if parte.get('polo') == 'PASSIVO':
            nome_passivo = parte.get('nome')
            doc_passivo = parte.get('documentosPrincipais', [{}])[0].get('numero')

    if data_distribuicao and (not data_inicial or data_distribuicao > data_inicial):
        return {
            'Data da distribuição': data_distribuicao.strftime('%d/%m/%Y'),
            'Polo Ativo': nome_ativo.title() if nome_ativo else 'Desconhecido',
            'CPF/CNPJ Ativo': formatar_documento(doc_ativo),
            'Polo Passivo': nome_passivo.title() if nome_ativo else 'Desconhecido',
            'CPF/CNPJ Passivo': formatar_documento(doc_passivo),
            'Nº do Processo': numero_processo,
            'Tribunal': estado_tribunal,
            'Valor Causa': locale.currency(valor_acao, grouping=True) if valor_acao else '',
            'Classe Judicial': classes,
            'Assunto': assunto,
        }
    return None


def processar_cnpj(cnpj, bearer_code, data_inicial, log_callback, processos_existentes):
    log_callback(f"📄 Iniciando análise para CNPJ: {cnpj}")
    proxima_pagina = False
    dados_dos_processos = []
    analisados = 0
    pagina_atual = 1

    while proxima_pagina or pagina_atual == 1:
        log_callback(f"🔄 Buscando página {pagina_atual} para CNPJ {cnpj}...")
        dados_pagina = api_jusbr(bearer_code, cnpj, proxima_pagina)
        if not dados_pagina:
            log_callback(f"⚠️ Página {pagina_atual} vazia. Encerrando para {cnpj}.")
            break
        processos = dados_pagina.get('content', [])
        if not processos:
            log_callback(f"⚠️ Nenhum processo retornado na página {pagina_atual}.")
            break
        analisados += len(processos)
        proxima = dados_pagina.get('searchAfter')
        proxima_pagina = f"{proxima[0]},{proxima[1]}" if proxima else None

        for processo in processos:
            numero_processo = processo.get('numeroProcesso')
            if numero_processo and numero_processo in processos_existentes:
                continue
            info = capturar_informacoes(processo, data_inicial)
            if info:
                salvar_informacoes_no_json(info)
                dados_dos_processos.append(info)

        if not proxima_pagina:
            log_callback("🚫 Não há mais páginas disponíveis.")
            break
        pagina_atual += 1

    log_callback(
        f"✅ Finalizado {cnpj}. Total capturado: {len(dados_dos_processos)} processos válidos em {pagina_atual - 1} páginas.")
    return len(dados_dos_processos)


def executar(data_inicial, log_callback, bearer_code, processos_existentes, lista_cpfs):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    cnpjs = [c.replace('.', '').replace('/', '').replace('-', '') for c in lista_cpfs]
    data_inicial = datetime.strptime(data_inicial, '%d/%m/%Y') if data_inicial else None
    total = 0
    for cnpj in cnpjs:
        total += processar_cnpj(cnpj, bearer_code, data_inicial, log_callback, processos_existentes)
    return total


def iniciar_driver(callback_token_encontrado):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("user-agent=Mozilla/5.0")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
    driver = webdriver.Chrome(service=Service(), options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
    })

    def monitorar_logs():
        while True:
            logs = driver.get_log("performance")
            for entry in logs:
                message = json.loads(entry["message"])["message"]
                if message["method"] == "Network.requestWillBeSent":
                    headers = message["params"]["request"]["headers"]
                    auth = headers.get("authorization") or headers.get("Authorization")
                    if auth:
                        callback_token_encontrado(auth)
                        return
            time.sleep(1)

    threading.Thread(target=monitorar_logs, daemon=True).start()
    return driver


class ConsultaJusbrApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta Jusbr - Validação de Processos")
        self.root.geometry("800x600")
        self.data_inicial = tk.StringVar()
        self.bearer_code = ''
        self.excel_selecionado = False
        self.caminho_excel = None
        self.processos_existentes = set()

        tk.Button(root, text="Abrir site do Jusbr", command=self.abrir_site_jusbr, bg="#337ab7", fg="white").pack(
            pady=(10, 20))
        tk.Label(root, text="Data Inicial (dd/mm/aaaa):").pack(pady=5)
        tk.Entry(root, textvariable=self.data_inicial).pack()
        tk.Button(root, text="Selecionar o Excel de Processos", command=self.selecionar_excel).pack(pady=5)
        tk.Button(root, text="Selecionar planilha de CPFs/CNPJs", command=self.selecionar_cpfs_cnpjs).pack(pady=5)

        self.iniciar_button = tk.Button(root, text="Iniciar Consulta", command=self.iniciar_consulta)
        self.iniciar_button.pack(pady=10)

        self.log_text = scrolledtext.ScrolledText(root, height=25, width=100, state='disabled')
        self.log_text.pack(pady=10)

    def abrir_site_jusbr(self):
        time.sleep(2)
        self.log("🌐 Iniciando navegador...")

        def token_encontrado(token):
            self.bearer_code = token
            self.log(f"✅ Token capturado.\nPode iniciar a consulta.")

        try:
            self.driver = iniciar_driver(token_encontrado)
            self.driver.get("https://portaldeservicos.pdpj.jus.br/")
            self.log("🔐 Faça login manualmente no site.")
            self.log("⏳ Aguardando token de autenticação...")
        except Exception as e:
            self.log(f"❌ Erro ao iniciar navegador: {e}")
            messagebox.showerror("Erro", str(e))

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def selecionar_excel(self):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not caminho_arquivo:
            return
        try:
            df = pd.read_excel(caminho_arquivo)
            if 'Nº do Processo' not in df.columns:
                messagebox.showerror("Erro", "A planilha deve conter a coluna 'Nº do Processo'")
                return
            self.processos_existentes = set(df['Nº do Processo'].dropna().astype(str))
            self.caminho_excel = caminho_arquivo
            self.excel_selecionado = True
            self.log(f"📁 Excel carregado. {len(self.processos_existentes)} processos existentes detectados.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar Excel: {e}")

    def selecionar_cpfs_cnpjs(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not caminho:
            return
        try:
            df = pd.read_excel(caminho)
            colunas_possiveis = ['CPF', 'CNPJ', 'CPF/CNPJ']
            for col in colunas_possiveis:
                if col in df.columns:
                    self.lista_cpfs_cnpjs = df[col].dropna().astype(str).tolist()
                    break
            else:
                messagebox.showerror("Erro", "A planilha deve conter uma coluna chamada 'CPF', 'CNPJ' ou 'CPF/CNPJ'")
                return
            self.log(f"📄 Lista de CPFs/CNPJs carregada com {len(self.lista_cpfs_cnpjs)} registros.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha de CPFs/CNPJs: {e}")

    def iniciar_consulta(self):
        data = self.data_inicial.get().strip()
        if data:
            try:
                datetime.strptime(data, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida (formato: dd/mm/aaaa)")
                return
        threading.Thread(target=self.executar_thread, args=(data,)).start()

    def executar_thread(self, data):
        try:
            self.log("🔍 Iniciando consulta...")
            if not self.bearer_code:
                self.log("❌ Token não capturado. Faça login antes.")
                return

            if not self.lista_cpfs_cnpjs:
                self.log("⚠️ Nenhuma planilha de CPFs/CNPJs carregada.")
                messagebox.showwarning("Atenção", "Você precisa selecionar a planilha de CPFs/CNPJs antes de iniciar.")
                return

            total = executar(data, self.log, self.bearer_code, self.processos_existentes, self.lista_cpfs_cnpjs)
            self.log(f"✅ Consulta finalizada com {total} resultados.")
            self.salvar_novos_processos()
        except Exception as e:
            self.log(f"❌ Erro: {e}")
            messagebox.showerror("Erro", str(e))

    def salvar_novos_processos(self):
        try:
            json_path = f'{NOME_ARQUIVO_PARA_SALVAR}.json'
            if not os.path.exists(json_path):
                self.log("📭 Nenhum novo processo encontrado. Nada a salvar.")
                return

            with open(json_path, 'r', encoding='utf-8') as f:
                novos_dados = json.load(f)
            df_novos = pd.DataFrame(novos_dados)

            if self.caminho_excel:
                df_existente = pd.read_excel(self.caminho_excel)
                df_final = pd.concat([df_existente, df_novos], ignore_index=True)
                df_final.to_excel(self.caminho_excel, index=False, engine='openpyxl')
                self.log("✅ Novos processos adicionados na planilha original.")
            else:
                novo_arquivo = f'{NOME_ARQUIVO_PARA_SALVAR}.xlsx'
                df_novos.to_excel(novo_arquivo, index=False, engine='openpyxl')
                self.log(f"📄 Planilha nova criada: {novo_arquivo}")

            os.remove(json_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar na planilha: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ConsultaJusbrApp(root)
    root.mainloop()
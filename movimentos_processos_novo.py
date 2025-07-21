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


def api_jusbr(log_callback, bearer_code, filtro, paginacao=None):
    time.sleep(0.4)
    url_base = 'https://portaldeservicos.pdpj.jus.br/api/v2/processos'
    headers = {'Authorization': f'{bearer_code}'}

    if len(filtro) == 20:
        query_params = {'numeroProcesso': filtro}
    else:
        query_params = {'cpfCnpjParte': filtro}

    if paginacao:
        query_params['searchAfter'] = paginacao
    response = requests.get(url_base, headers=headers, params=query_params)

    if response.status_code != 200:
        if log_callback:
            log_callback(f"❌ Erro ao buscar processos (status {response.status_code}): {response.text}")
        return None

    return response.json()


def formatar_documento(documento):
    if documento and isinstance(documento, str):
        documento = documento.strip().replace('.', '').replace('-', '').replace('/', '')
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


def capturar_informacoes(processo):
    tramitacoes = processo.get('tramitacoes', [])
    tramitacao = tramitacoes[0] if tramitacoes else {}
    partes = tramitacao.get('partes', [])
    classes = tramitacao.get('classe', '') or ''
    assunto_info = tramitacao.get('assunto', [{}])
    assunto = assunto_info[0].get("descricao") if assunto_info and isinstance(assunto_info, list) else ''

    distribuicao = tramitacao.get('dataHoraUltimaDistribuicao') or ''
    try:
        data_distribuicao = datetime.strptime(distribuicao.split('.')[0], '%Y-%m-%dT%H:%M:%S') if distribuicao else None
    except Exception:
        data_distribuicao = None

    valor_acao = tramitacao.get('valorAcao') or 0
    estado_tribunal = processo.get('siglaTribunal', '') or ''
    numero_processo = processo.get('numeroProcesso', '') or ''

    nome_ativo, doc_ativo, nome_passivo, doc_passivo = None, None, None, None
    for parte in partes:
        polo = parte.get('polo', '')
        if polo == 'ATIVO':
            nome_ativo = parte.get('nome')
            doc_ativo = (parte.get('documentosPrincipais') or [{}])[0].get('numero')
        elif polo == 'PASSIVO':
            nome_passivo = parte.get('nome')
            doc_passivo = (parte.get('documentosPrincipais') or [{}])[0].get('numero')

    return {
        'Data da distribuição': data_distribuicao.strftime('%d/%m/%Y') if data_distribuicao else '',
        'Polo Ativo': nome_ativo.title() if nome_ativo else 'Desconhecido',
        'CPF/CNPJ Ativo': formatar_documento(doc_ativo) or '',
        'Polo Passivo': nome_passivo.title() if nome_passivo else 'Desconhecido',
        'CPF/CNPJ Passivo': formatar_documento(doc_passivo) or '',
        'Nº do Processo': numero_processo,
        'Tribunal': estado_tribunal,
        'Valor Causa': locale.currency(valor_acao, grouping=True) if valor_acao else '',
        'Classe Judicial': classes or '',
        'Assunto': assunto or '',
    }


def processar_numero(numero, bearer_code, data_inicial, data_final, log_callback, movimentos_existentes):
    tipo = (
        "CPF" if len(numero) == 11 else
        "CNPJ" if len(numero) == 14 else
        "Número do Processo" if len(numero) == 20 else
        "Número Desconhecido"
    )
    log_callback(f"📄 Iniciando análise para {tipo}: {numero}")

    proxima_pagina = False
    dados_dos_processos = []
    analisados = 0
    pagina_atual = 1
    total_processos = None

    while proxima_pagina or pagina_atual == 1:
        log_callback(f"🔄 Buscando página {pagina_atual} para {tipo} {numero}...")
        dados_pagina = api_jusbr(log_callback, bearer_code, numero, proxima_pagina)
        if not dados_pagina:
            log_callback(f"⚠️ Página {pagina_atual} vazia. Encerrando para {numero}.")
            break

        if total_processos is None:
            total_processos = dados_pagina.get('total', 0)
            log_callback(f"📦 Total de processos encontrados: {total_processos}")

        processos = dados_pagina.get('content', [])
        if not processos:
            log_callback(f"⚠️ Nenhum processo retornado na página {pagina_atual}.")
            break

        analisados += len(processos)
        proxima = dados_pagina.get('searchAfter')
        proxima_pagina = f"{proxima[0]},{proxima[1]}" if proxima else None

        for processo in processos:
            numero_processo = processo.get('numeroProcesso')

            processo_info = capturar_informacoes(processo)
            movimentos_info = obter_movimentos(log_callback, bearer_code, numero_processo)

            if movimentos_info:
                for movimento in movimentos_info:
                    data_mov = movimento.get('data')

                    if data_inicial and data_mov < data_inicial:
                        continue
                    if data_final and data_mov > data_final:
                        continue

                    if movimentos_existentes:
                        ultimo_registro = movimentos_existentes.get(numero_processo)

                        if ultimo_registro and data_mov <= datetime.strptime(ultimo_registro, '%d/%m/%Y %H:%M:%S'):
                            continue

                    if processo_info:
                        processo_info['Data do Movimento'] = data_mov.strftime('%d/%m/%Y') if data_mov else ''
                        processo_info['Movimento'] = movimento.get('movimento')
                        processo_info['DataHora'] = data_mov.strftime('%d/%m/%Y %H:%M:%S') if data_mov else ''
                        salvar_informacoes_no_json(processo_info)
                        dados_dos_processos.append(processo_info)

        if analisados >= total_processos:
            log_callback("🚫 Não há mais páginas disponíveis ou todos os processos foram analisados.")
            break

        pagina_atual += 1

    log_callback(
        f"✅ Finalizado {numero}. Total capturado: {len(dados_dos_processos)} movimentos válidos.")
    return len(dados_dos_processos)


def executar(data_inicial, data_final, log_callback, bearer_code, movimentos_existentes, para_capturar):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    para_capturar_limpas = [c.replace('.', '').replace('/', '').replace('-', '') for c in para_capturar if c]

    data_inicial = datetime.strptime(data_inicial, '%d/%m/%Y') if data_inicial else None
    data_final = datetime.strptime(data_final, '%d/%m/%Y') if data_final else None
    total = 0
    for numero in para_capturar_limpas:
        total += processar_numero(numero, bearer_code, data_inicial, data_final, log_callback, movimentos_existentes)

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


def obter_movimentos(log_callback, bearer_code, numero_processo):
    time.sleep(0.4)
    url = f'https://portaldeservicos.pdpj.jus.br/api/v2/processos/{numero_processo}'
    headers = {'Authorization': bearer_code}
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        if log_callback:
            log_callback(
                f"❌ Erro ao obter movimentos do processo {numero_processo} (status {response.status_code}): {response.text}")
        return []
    data = response.json()
    movimentos = []
    try:
        tramitacao = data[0].get('tramitacaoAtual', [])
        for movimento in tramitacao.get('movimentos', []):
            data_mov = movimento.get('dataHora', '')
            descricao = movimento.get('descricao', '')
            if data_mov and descricao:
                try:
                    data_formatada = datetime.strptime(data_mov.split('.')[0], '%Y-%m-%dT%H:%M:%S')
                    movimentos.append({
                        'data': data_formatada,
                        'movimento': descricao
                    })
                except Exception:
                    continue
    except Exception:
        pass
    return movimentos


class ConsultaJusbrApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta Jusbr - Movimentos dos Processos")
        self.root.geometry("800x600")
        self.data_inicial = tk.StringVar()
        self.data_final = tk.StringVar()
        self.bearer_code = 'Bearer eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICI1dnJEZ1hCS21FLTdFb3J2a0U1TXU5VmxJZF9JU2dsMnY3QWYyM25EdkRVIn0.eyJleHAiOjE3NTMxNDQyOTAsImlhdCI6MTc1MzExNTQ5MCwiYXV0aF90aW1lIjoxNzUzMTE1NDg1LCJqdGkiOiI1ZDkzNGI0ZS01MzgwLTQ0N2YtYjY2Ni02NjE5NjZjMmIwODEiLCJpc3MiOiJodHRwczovL3Nzby5jbG91ZC5wamUuanVzLmJyL2F1dGgvcmVhbG1zL3BqZSIsImF1ZCI6WyJicm9rZXIiLCJhY2NvdW50Il0sInN1YiI6IjhkMGMzYmNjLTNkOWItNGZlMy04ZThjLWFhN2M0Mzk5NGEwYiIsInR5cCI6IkJlYXJlciIsImF6cCI6InBvcnRhbGV4dGVybm8tZnJvbnRlbmQiLCJub25jZSI6IjkzYWE2M2E4LTUxMzYtNGY3Mi1iYWJmLTg3MjcwZTI3NjVmNSIsInNlc3Npb25fc3RhdGUiOiIwYzkyZTM0Yy05N2NhLTQxMTgtOTNmMi1hNjBiZjdiM2JlZDIiLCJhY3IiOiIwIiwiYWxsb3dlZC1vcmlnaW5zIjpbImh0dHBzOi8vcG9ydGFsZGVzZXJ2aWNvcy5wZHBqLmp1cy5iciJdLCJyZWFsbV9hY2Nlc3MiOnsicm9sZXMiOlsiZGVmYXVsdC1yb2xlcy1wamUiLCJvZmZsaW5lX2FjY2VzcyIsInVtYV9hdXRob3JpemF0aW9uIl19LCJyZXNvdXJjZV9hY2Nlc3MiOnsiYnJva2VyIjp7InJvbGVzIjpbInJlYWQtdG9rZW4iXX0sImFjY291bnQiOnsicm9sZXMiOlsibWFuYWdlLWFjY291bnQiLCJtYW5hZ2UtYWNjb3VudC1saW5rcyIsInZpZXctcHJvZmlsZSJdfX0sInNjb3BlIjoib3BlbmlkIHByb2ZpbGUgZW1haWwiLCJzaWQiOiIwYzkyZTM0Yy05N2NhLTQxMTgtOTNmMi1hNjBiZjdiM2JlZDIiLCJBY2Vzc2FSZXBvc2l0b3JpbyI6Ik9rIiwiZW1haWxfdmVyaWZpZWQiOnRydWUsIm5hbWUiOiJFRFVBUkRPIFBFUkVJUkEgR09NRVMiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiI3Njc4NzA0NDAyMCIsImdpdmVuX25hbWUiOiJFRFVBUkRPIFBFUkVJUkEiLCJmYW1pbHlfbmFtZSI6IkdPTUVTIiwiY29ycG9yYXRpdm8iOlt7InNlcV91c3VhcmlvIjo1MzQ3MjA5LCJub21fdXN1YXJpbyI6IkVEVUFSRE8gUEVSRUlSQSBHT01FUyIsIm51bV9jcGYiOiI3Njc4NzA0NDAyMCIsInNpZ190aXBvX2NhcmdvIjoiQURWIiwiZmxnX2F0aXZvIjoiUyIsInNlcV9zaXN0ZW1hIjowLCJzZXFfcGVyZmlsIjowLCJkc2Nfb3JnYW8iOiJPQUIiLCJzZXFfdHJpYnVuYWxfcGFpIjowLCJkc2NfZW1haWwiOiJzZWNyZXRhcmlhQGVkdWFyZG9nb21lcy5hZHYuYnIiLCJzZXFfb3JnYW9fZXh0ZXJubyI6MCwiZHNjX29yZ2FvX2V4dGVybm8iOiJPQUIiLCJvYWIiOiJSUzkxNjMxIn1dLCJlbWFpbCI6InNlY3JldGFyaWFAZWR1YXJkb2dvbWVzLmFkdi5iciJ9.JdDcxHhOKoZHsDH76lyj09mR5VgM8t0AWU0gjIJjFVamY35TtUC2utmQDGthVyws4FYvcTWhQEbbI12AIEpMHbENC6Gp-b2V0-TlFQUyP9vAOctRbOkbo9sECzc4fETRQlrYSRfTOLebB0dRv3ZPqQGypGIUlxIRbz5x-tiOjBERERE9N09CLm9fTs_KKqiBibzy0aFRHWhhgLPkCDU1dDrOUp1nFu10zrRhywDtJ-Bt3qgCe9ZO1lTUPwtbDyPsRt4ULOQFjmBpi6E5mAFxFf0JINmDhAaRjbrAKnJn3r-JsoxAPObH7p6rEmK_6K7jL4cTs3q-S-nuCT8pQ3SvmA'
        self.caminho_excel = None
        self.movimentos_existentes = {}

        tk.Button(root, text="Abrir site do Jusbr", command=self.abrir_site_jusbr, bg="#337ab7", fg="white").pack(
            pady=(10, 20))
        tk.Label(root, text="Data Inicial (dd/mm/aaaa):").pack(pady=5)
        tk.Entry(root, textvariable=self.data_inicial).pack()
        tk.Label(root, text="Data Final (opcional - dd/mm/aaaa):").pack(pady=5)
        tk.Entry(root, textvariable=self.data_final).pack()
        tk.Button(root, text="Selecionar o Excel de Processos", command=self.selecionar_excel).pack(pady=5)
        tk.Button(root, text="Selecionar o Excel dos CPFs/CNPJs/Processos", command=self.selecionar_cpfs_cnpjs).pack(
            pady=5)

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

            colunas_necessarias = ['Nº do Processo', 'DataHora']
            colunas_ausentes = [col for col in colunas_necessarias if col not in df.columns]
            if colunas_ausentes:
                mensagem = "A planilha deve conter as colunas obrigatórias: " + ", ".join(colunas_ausentes)
                messagebox.showerror("Erro", mensagem)
                return

            df = df.dropna(subset=['Nº do Processo', 'DataHora'])

            for processo, grupo in df.groupby('Nº do Processo'):
                ultima_linha = grupo.sort_values(by='DataHora', ascending=False).iloc[0]
                self.movimentos_existentes[processo] = ultima_linha['DataHora']

            self.caminho_excel = caminho_arquivo
            self.log(f"📁 Excel carregado. {len(self.movimentos_existentes)} movimentos existentes detectados.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar Excel: {e}")

    def selecionar_cpfs_cnpjs(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not caminho:
            return
        try:
            df = pd.read_excel(caminho)
            colunas_possiveis = ['CPF', 'CNPJ', 'CPF/CNPJ', 'Nº do Processo']
            for col in colunas_possiveis:
                if col in df.columns:
                    self.para_capturar = df[col].dropna().astype(str).tolist()
                    break
            else:
                messagebox.showerror("Erro",
                                     "A planilha deve conter pelo menos uma das seguintes colunas: 'CPF', 'CNPJ', 'CPF/CNPJ' ou 'Nº do Processo'")
                return
            self.log(f"📄 Lista de CPFs/CNPJs/Processos carregada com {len(self.para_capturar)} registros.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha de CPFs/CNPJs/Processos: {e}")

    def iniciar_consulta(self):
        data_ini = self.data_inicial.get().strip()
        data_fim = self.data_final.get().strip()
        if not data_ini:
            messagebox.showerror("Erro", "A Data Inicial é obrigatória.")
            return
        try:
            datetime.strptime(data_ini, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Data Inicial inválida (formato: dd/mm/aaaa)")
            return
        if data_fim:
            try:
                datetime.strptime(data_fim, "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Erro", "Data Final inválida (formato: dd/mm/aaaa)")
                return
        threading.Thread(target=self.executar_thread, args=(data_ini, data_fim)).start()
        return

    def iniciar_consulta_backup(self):
        data = self.data_inicial.get().strip()
        if data:
            try:
                datetime.strptime(data, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida (formato: dd/mm/aaaa)")
                return
        threading.Thread(target=self.executar_thread, args=(data,)).start()

    def executar_thread(self, data_ini, data_fim):
        try:
            self.log("🔍 Iniciando consulta...")

            if not self.bearer_code:
                self.log("❌ Token não capturado. Faça login antes.")
                return

            if not hasattr(self, 'para_capturar') or not self.para_capturar:
                self.log("⚠️ A planilha de CPFs/CNPJs/Processos não foi selecionada.")
                messagebox.showwarning("Atenção",
                                       "Você precisa selecionar a planilha de CPFs/CNPJs/Processos antes de iniciar.")
                return

            total = executar(data_ini, data_fim, self.log, self.bearer_code, self.movimentos_existentes,
                             self.para_capturar)
            self.log(f"✅ Consulta finalizada com {total} resultados.")
            self.salvar_novos_processos()

        except Exception as e:
            self.log(f"❌ Erro: {e}")
            messagebox.showerror("Erro", str(e))

    def salvar_novos_processos(self):
        try:
            data = datetime.now().strftime('%Y-%m-%d %H %M')

            json_path = f'{NOME_ARQUIVO_PARA_SALVAR}.json'
            if not os.path.exists(json_path):
                self.log("📭 Nenhum novo processo encontrado. Nada a salvar.")
                return

            with open(json_path, 'r', encoding='utf-8') as f:
                novos_dados = json.load(f)
            df_novos = pd.DataFrame(novos_dados)

            novo_arquivo = f'{NOME_ARQUIVO_PARA_SALVAR} - {data}.xlsx'
            df_novos.to_excel(novo_arquivo, index=False, engine='openpyxl')
            self.log(f"📄 Planilha nova criada: {novo_arquivo}")

            os.remove(json_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar na planilha: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ConsultaJusbrApp(root)
    root.mainloop()
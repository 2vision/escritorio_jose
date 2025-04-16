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


NOME_ARQUIVO_PARA_SALVAR = 'Processos Verificados'
MOVIMENTOS = [
    "AÇÃO ANULATÓRIA AJUIZADA",
    "AÇÃO ARRESTO AJUIZADA",
    "AÇÃO CAUTELAR AJUIZADA",
    "AÇÃO RESCISÓRIA AJUIZADA",
    "AÇÃO SEQÜESTRO DE BENS AJUIZADA",
    "ACÓRDÃO DESFAVORÁVEL À CAIXA",
    "ACÓRDÃO FAVORÁVEL À CAIXA",
    "ACÓRDÃO PARCIALMENTE FAVORÁVEL",
    "ACORDO FIRMADO",
    "ACORDO NÃO CUMPRIDO",
    "ACORDO PARCIAL FIRMADO",
    "ACORDO PETIÇÃO",
    "ACORDO PROPOSTA CAIXA",
    "ACORDO PROPOSTA CAIXA REJEITADA",
    "ADJUDICAÇÃO/ARREMATAÇÃO PELA CAIXA",
    "ADJUDICAÇÃO/ARREMATAÇÃO POR TERCEIRO",
    "AGRAVO DE INSTRUMENTO RESPOSTA",
    "AGRAVO DE PETIÇÃO INTERPOSTO",
    "AGRAVO DE PETIÇÃO RESPOSTA",
    "AGRAVO EM RECURSO ESPECIAL INTERPOSTO",
    "AGRAVO EM RECURSO ESPECIAL RESPOSTA",
    "AGRAVO EM RECURSO EXTRAORDINARIO INTERPOSTO",
    "AGRAVO EM RECURSO EXTRAORDINARIO RESPOSTA",
    "AGRAVO INOMINADO/INTERNO (REGIMENTAL) RESPOST",
    "AGRAVO INTERNO",
    "AGRAVO RETIDO INTERPOSTO",
    "AGRAVO RETIDO RESPOSTA",
    "ALVARÁ DE LEVANTAMENTO PARA A CAIXA RETIRADO",
    "ANÁLISE DE AUTOS",
    "ARQUIVAMENTO",
    "ASSEMBLEIA DE CREDORES - COMPARECIMENTO",
    "ASSEMBLEIA DE CREDORES - INTIMACAO",
    "ATO PROCESSUAL TERCEIRIZADO",
    "AUDIÊNCIA DE CONC. OU INAUGURAL - ADV. CRED.",
    "AUDIÊNCIA DE INSTRUÇÃO E JULG. - ADV. CRED.",
    "AUDIÊNCIA DE JUSTIFICAÇÃO",
    "AUDIÊNCIA INSTITUCIONAL",
    "AVALIAÇÃO DE BENS",
    "BAIXA A ORIG. PARA UNIF. DE JURISTPRUDENCIA",
    "BAIXA DOS AUTOS À ORIGEM",
    "BLOQUEIO CONSULTA AUTOMÁTICA/BOLETO",
    "BOLETO EMITIDO",
    "BOLETO/ PAGAMENTO TOTAL",
    "BOLETO/PAGAMENTO PARCIAL",
    "BUSCA E APREENSÃO DO BEM EFETIVADA",
    "CÁLCULO LIQUIDAÇÃO - ATENDIMENTO",
    "CÁLCULO LIQUIDAÇÃO - REQUERIMENTO",
    "CARGA - DEVOLUÇÃO DOS AUTOS",
    "CARGA - RETIRADA DOS AUTOS",
    "CARTA DE ARREMATAÇÃO/ADJUDICAÇÃO - EXPEDIÇÃO",
    "CARTA DE ARREMATAÇÃO/ADJUDICAÇÃO - REGISTRO",
    "CARTA DE ORDEM - DEVOLUÇÃO",
    "CARTA DE ORDEM - DISTRIBUIÇÃO",
    "CARTA DE ORDEM - EXPEDIÇÃO",
    "CARTA DE SENTENÇA",
    "CARTA PRECATÓRIA - DEVOLUÇÃO",
    "CARTA PRECATÓRIA - DISTRIBUIÇÃO",
    "CARTA PRECATÓRIA - EXPEDIÇÃO",
    "CITAÇÃO/NOTIFICAÇÃO",
    "COMUNICAÇÃO À GEATS/GETEN",
    "COMUNICAÇÃO AO JURIR DE ORIGEM",
    "COMUNICAÇÃO AO JURIR DE TRF",
    "COMUNICAÇÃO DECISÃO JUDICIAL À ÁREA GESTORA",
    "CONCILIACAO EXTRAJUDICIAL - DEVOL P ADEQUACAO",
    "CONCILIACAO EXTRAJUDICIAL - NAO OFERTADA",
    "CONCILIACAO EXTRAJUDICIAL - TRANSF DE JURIR",
    "CONCILIACAO EXTRAJUDICIAL APROVADA",
    "CONCILIACAO EXTRAJUDICIAL NAO APROVADA",
    "CONCILIACAO EXTRAJUDICIAL-NAO ACEITA",
    "CONCILIAÇÃO EXTRAJUDICIAL-PROC PRÉ EXISTENTE",
    "CONCLUSO PARA JULGAMENTO",
    "CONSULTA RECEBIDA SOBRE PROCESSO JUDICIAL",
    "CONSULTA RESPONDIDA SOBRE PROCESSO JUDICIAL",
    "CONTESTAÇÃO OU DEFESA TRABALHISTA",
    "CONVERSÃO DE AGRAVO - DECISÃO MONOCRÁTICA",
    "CONVERSÃO DE ARRESTO EM PENHORA",
    "CONVERSÃO EM EXECUÇÃO",
    "CRÉDITO LIQUIDAÇÃO DO FGTS - ATENDIMENTO",
    "CRÉDITO LIQUIDAÇÃO DO FGTS - REQUERIMENTO",
    "CRÉDITO PRESCRITO IMPOSSIBILIDADE DE AJUIZAME",
    "CUMPRIMENTO ESPONTÂNEO DA OBRIGAÇÃO",
    "CUMPRIMENTO ESPONTÂNEO DE OBRIGAÇÃO ADMIN",
    "CUMPRIMENTO SENTENÇA - REQUERIMENTO DO CREDOR",
    "DECISÃO AFETAÇÃO/INSTAURAÇÃO DE IRDR",
    "DECISÃO EM EXECUÇÃO DE JULGADO",
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
    "DECRETO DE PRISÃO",
    "DEMONSTRATIVO DÉBITO/MEMÓRIA CÁLCULOS PETIÇÃO",
    "DEP GARANTIA JUIZO - OUTROS BANCOS",
    "DESARQUIVAMENTO PETIÇÃO",
    "DESBLOQUEIO CONSULTA AUTOMÁTICA/BOLETO",
    "DESERÇÃO",
    "DESISTÊNCIA DA AÇÃO PELA CAIXA PETIÇÃO",
    "DESISTÊNCIA DE RECURSO PELA CAIXA PETIÇÃO",
    "DESISTÊNCIA/RENÚNCIA PELA PARTE CONTRÁRIA",
    "DESPACHO DE MERO EXPEDIENTE INTIMAÇÃO",
    "DESPACHO SANEADOR/JULG. ANTECIPADO INTIMAÇÃO",
    "DEVOLUCAO DE HONORARIOS DE SUCUMBENCIA",
    "DEVOLVIDO UNIDADE ORIGEM PARA COMPLEMENTAÇÃO",
    "DILAÇÃO OU RESTITUIÇÃO DE PRAZO PETIÇÃO",
    "DILIGÊNCIA ADMINISTRATIVA EXPEDIDA",
    "DILIGÊNCIA ADMINISTRATIVA EXPEDIDA - EMGEA",
    "DILIGÊNCIA ADMINISTRATIVA RECEBIDA",
    "DILIGÊNCIA ADMINISTRATIVA RECEBIDA - EMGEA",
    "DILIGENCIA TERCEIRIZADA",
    "DISTRIBUIÇÃO",
    "EDITAL DE CITAÇÃO REQUERIDO",
    "EDITAL DE PRAÇA REQUERIDO",
    "EM ANALISE- CONCILIACAO EXTRAJUDICIAL",
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
    "FALHA PRODUTO/SERVIÇO",
    "FGTS - COMPROVAÇÃO DE DEPÓSITO",
    "FRAUDE À EXECUÇÃO PETIÇÃO",
    "HONORÁRIOS - PARCELAMENTO AUTORIZADO PELA ADV",
    "HONORÁRIOS ADVOCATÍCIOS EXECUÇÃO NOS AUTOS",
    "IMISSÃO OU REINTEGRAÇÃO DE POSSE EFETIVADA",
    "IMPOSSIBILIDADE DE EXTINÇÃO",
    "IMPUGNAÇÃO À ASSISTÊNCIA JUDICIÁRIA GRATUITA",
    "IMPUGNAÇÃO AO VALOR DA CAUSA",
    "IMPUGNAÇÃO AO VALOR DA CAUSA RÉPLICA",
    "IMPUGNAÇÃO AOS CÁLCULOS",
    "INCID DE UNIFORMIZ DE JURISPRUD - RESPOSTA",
    "INCIDENTE DE FALSIDADE",
    "INCIDENTE DE UNIFORMIZAÇÃO DE JURISPRUDÊNCIA",
    "INFOJUD",
    "INFORMACAO DE CITACAO/NOTIFICACAO AO GESTOR",
    "INSTAURAÇÃO DO INQUÉRITO POLICIAL",
    "INTERROGATÓRIO DO ACUSADO",
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
    "JUNTADA DE SUBSTABELECIMENTO PETIÇÃO",
    "LIMINAR - COMUNICACAO A AREA GESTORA",
    "LIMINAR CONCEDIDA",
    "LIMINAR NEGADA/REVOGADA/CASSADA",
    "LIMINAR PARCIALMENTE CONCEDIDA",
    "MANDADO DE IMISSÃO DE POSSE",
    "MANDADO DE SEGURANCA INTERPOSTO",
    "MANIFESTAÇÃO PROCESSUAL - OUTRAS",
    "NÃO RECEBIMENTO DA DENÚNCIA",
    "NOTIFICAÇÃO EXTRAJUDICIAL",
    "OFERECIMENTO DA DENÚNCIA",
    "PAGAMENTO ALUGUEL - DEFERIDO",
    "PAGAMENTO PARCIAL DA CONDENAÇÃO",
    "PEDIDO DE ANTECIP. TUTELA/LIMINAR IMPUGNAÇÃO",
    "PEDIDO DE REEXAME",
    "PEDIDO DE RESSARCIMENTO - CAIXAPAR",
    "PEDIDO DE RESSARCIMENTO - PANAMERICANO",
    "PEDIDO DE RESSARCIMENTO - SUL AMÉRICA",
    "PEDIDO DE SUBSÍDIOS/CUMPRIMENTO - NÃO ATENDIM",
    "PENHORA - BAIXA",
    "PENHORA - GARANTIA DO JUÍZO",
    "PENHORA - INTIMAÇÃO",
    "PENHORA - LAVRATURA",
    "PENHORA - REGISTRO/AVERBAÇÃO NO OFÍCIO IMOBIL",
    "PERÍCIA - FORMULAÇÃO DE QUESITOS",
    "PERÍCIA - MANIFESTAÇÃO",
    "PESQUISA DE BENS NEGATIVA",
    "PESQUISA DE BENS SEM ATENDIMENTO",
    "PETIÇÃO DA PARTE CONTRÁRIA-LITISCONSORCIAL",
    "PETIÇÃO HABILITAÇÃO/DIVERGÊNCIA DE CRÉDITO",
    "PJE",
    "POUPANÇA - CÁLCULO LIQUIDAÇÃO - ATENDIMENTO",
    "POUPANÇA - CÁLCULO LIQUIDAÇÃO - REQUERIMENTO",
    "POUPANÇA - EXTRATO - RECEBIMENTO",
    "POUPANÇA - EXTRATO - REQUERIMENTO",
    "POUPANÇA - LOCALIZAÇÃO CONTA - ATENDIMENTO",
    "POUPANÇA - LOCALIZAÇÃO CONTA - REQUERIMENTO",
    "POUPANCA - PRESCRICAO - EXECUCAO ACP",
    "POUPANCA - PRESCRICAO - PLANO BRESSER",
    "POUPANCA - PRESCRICAO - PLANO COLLOR I",
    "POUPANCA - PRESCRICAO - PLANO COLLOR II",
    "POUPANCA - PRESCRICAO - PLANO VERAO",
    "POUPANCA-CALCULO DE LIQUIDACAO-DESNECESSARIO",
    "PRAÇA OU LEILÃO ARREMATAÇÃO PELA CAIXA",
    "PRAÇA OU LEILÃO CANCELADO OU SEM LICITANTE",
    "PRAÇA OU LEILÃO INTIMAÇÃO",
    "PROCESSO APTO PARA ACORDO",
    "PROCESSO INAPTO PARA ACORDO/DEPURADO",
    "PRODUÇÃO DE PROVAS PETIÇÃO",
    "PROJ DE DESISTÊNCIA - TRF - IMPOSSIBILIDADE",
    "PROJETO DE DESISTÊNCIA - TRF - POSSIBILIDADE",
    "PROTOCOLO ELETRÔNICO",
    "RAZÕES FINAIS OU MEMORIAIS PETIÇÃO",
    "REATIVAÇÃO DE PROCESSO EXTINTO",
    "RECEBIMENTO DA DENÚNCIA",
    "RECONVENÇÃO",
    "RECUPERACAO DE VALORES - FEITOS DIVERSOS",
    "RECUPERACAO JUDICIAL - ENCERRAMENTO",
    "RECUPERACAO JUDICIAL - INDEFERIMENTO",
    "RECUPERACAO JUDICIAL - LIQUIDACAO DO CREDITO",
    "RECUPERACAO JUDICIAL - PLANO - ADITAMENTO",
    "RECUPERACAO JUDICIAL - PLANO - APROVACAO AGC",
    "RECUPERACAO JUDICIAL - PLANO - HOMOLOGACAO",
    "RECUPERACAO JUDICIAL - PLANO - INTIMACAO",
    "RECUPERACAO JUDICIAL - PLANO - OBJECAO",
    "RECUPERACAO JUDICIAL - PLANO - REPROVACAO AGC",
    "RECUPERACAO JUDICIAL - REQUERIMENTO",
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
    "REDISTRIBUICAO DE FORO",
    "REMESSA À PROCURADORIA DA REPÚBLICA",
    "REMESSA À TURMA RECURSAL",
    "REMESSA À UNIDADE DESCENTRALIZADA",
    "REMESSA AO TAC - TRIBUNAL DE ALÇADA CÍVEL",
    "REMESSA AO TJE - TRIBUNAL DE JUSTIÇA ESTADUAL",
    "REMESSA AO TRF",
    "REMESSA AO TRT",
    "REMESSA AO TST",
    "REMESSA DE EXPEDIENTE À POLÍCIA FEDERAL",
    "REMIÇÃO",
    "RENAJUD",
    "RENÚNCIA AO PRAZO RECURSAL PETIÇÃO",
    "RENÚNCIA DIREITO DE AÇÃO PELA CAIXA PETIÇÃO",
    "RÉPLICA",
    "RÉPLICA À IMPUGNAÇÃO",
    "RESPONSABILIDADE ANALISADA",
    "RESSARCIMENTO - PEDIDO ENCAMINHADO",
    "REUNIÃO PRÉVIA COM PREPOSTO",
    "REVISÃO CRIMINAL",
    "SALDO EM CONTA JUDICIAL NÃO LEVANTADO",
    "SENTENÇA DESFAVORÁVEL À CAIXA",
    "SENTENÇA EXTINÇÃO POR CUMPRIMENTO OBRIGAÇÃO",
    "SENTENÇA EXTINÇÃO SEM RESOLUÇÃO DO MÉRITO",
    "SENTENÇA FAVORÁVEL À CAIXA",
    "SENTENÇA HOMOLOGATÓRIA DE ACORDO",
    "SENTENÇA HOMOLOGATÓRIA NA EXECUÇÃO",
    "SENTENÇA INCOMPETÊNCIA DE FORO",
    "SENTENÇA PARCIALMENTE FAVORÁVEL",
    "SH - DEFERIDO INGRESSO",
    "SH - PEDIDO DE INGRESSO",
    "SH - TRÂNSITO - DEFERIDO O INGRESSO",
    "SH - TRÂNSITO - INDEFERIDO O INGRESSO",
    "SH-SUBSÍDIOS RECEBIDOS PARA INGRESSO",
    "SINAD - REINCLUSÃO INDEVIDA - FALHA SISTEMA",
    "SOBRESTAMENTO - REC. REPETITIVO/STJ",
    "SOBRESTAMENTO - REPERC. GERAL/STF",
    "SOLICIT. DE ALVARÁ DE LEVANTAMENTO PARA CAIXA",
    "SOLICITAÇÃO DE AUTORIZAÇÃO DE PAGAMENTO",
    "SOLICITAÇÃO DE CÓPIAS",
    "SOLICITAÇÃO DE PAGAMENTO-DEPÓSITO",
    "SOLICITAÇÃO DE TESTEMUNHAS",
    "SOLICITAÇÃO PREPOSTO CAIXA",
    "SOLICITAÇÃO PREPOSTO ESPECIALIZADO",
    "SUBSIDIOS AUSENTES- DEFESA GENERICA",
    "SUBSÍDIOS RECEBIDOS",
    "SUBSÍDIOS SOLICITADOS",
    "SUMULA DE CONCILIACAO EXTRAJUDICIAL - SCE",
    "SUSPENSÃO COM ARQUIVAMENTO SEM BAIXA",
    "SUSPENSÃO DE SEGURANÇA INTERPOSTA",
    "SUSPENSÃO DO PROCESSO PETIÇÃO",
    "SUSTENTAÇÃO ORAL REALIZADA",
    "TAC - TERMO DE AJUSTE DE CONDUTA",
    "TERMO ADESÃO FGTS - HOMOLOGAÇÃO",
    "TERMO ADESÃO FGTS - HOMOLOGAÇÃO NEGADA",
    "TERMO ADESÃO FGTS - RECEBIMENTO",
    "TERMO ADESÃO FGTS - REQUERIMENTO",
    "TRANSITO EM JULGADO A CLASSIFICAR",
    "TRÂNSITO JULGADO DECISÃO DESFAVORÁVEL À CAIXA",
    "TRÂNSITO JULGADO DECISÃO FAVORÁVEL À CAIXA",
    "TRÂNSITO JULGADO DECISÃO PARCIALM. FAVORÁVEL",
    "TUTELA ANTECIPADA CONCEDIDA",
    "TUTELA ANTECIPADA NEGADA/REVOGADA/CASSADA",
    "TUTELA ANTECIPADA PARCIALMENTE CONCEDIDA",
    "VRE - RESPOSTA DE CÁLCULO",
    "VRE - SOLICITAÇÃO AUXÍLIO GETEN",
    "VRE - SOLICITAÇÃO DE CÁLCULO",
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


processos = dados_planilha()
cookies = pegar_cookies()
for processo in processos:
    numeros_expedientes, areas_judiciais  = consulta_numero_expediente(cookies, processo)
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
        'numero_expediente': numero_expediente,
        'area_judicial': areas_judiciais[0],
        'movimento_dia_seguinte': movimento_principal_dia_seguinte
    }
    salvar_informacoes_no_json(processo_atualizado, NOME_ARQUIVO_PARA_SALVAR)

salvar_informacoes_no_excel()

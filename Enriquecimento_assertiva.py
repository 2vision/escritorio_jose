from urllib.parse import quote
import pandas as pd
import os
from openpyxl.reader.excel import load_workbook

# Listar todos os arquivos no diretório
diretorio = 'Enriquecer'
arquivos_no_diretorio = os.listdir(diretorio)

# Função para validar e formatar o telefone
def validar_telefone(telefone):
    telefone = telefone.replace('+', '')  # Remove o prefixo '+'
    if len(telefone) < 12:
        return ''
    codigo_pais = telefone[:2]
    ddd = telefone[2:4]
    numero = telefone[4:]

    if int(ddd) <= 28:
        if len(numero) == 8:
            numero = '9' + numero
        elif len(numero) == 9 and numero[0] != '9':
            numero = '9' + numero[1:]
    else:
        if len(numero) == 9 and numero[0] == '9':
            numero = numero[1:]

    return f'+{codigo_pais}{ddd}{numero}'

def replicar_linhas(planilha, indice_linha, socios_com_telefone, emails, is_cnpj):
    if not socios_com_telefone:
        return planilha

    socios_unicos = socios_com_telefone[:3]

    if len(socios_unicos) > 0:
        primeiro_socio = socios_unicos[0]
        telefone_validado = validar_telefone(primeiro_socio["Telefone"])

        if "Telefone" not in planilha.columns:
            planilha["Telefone"] = ""

        planilha.at[indice_linha, "Telefone"] = telefone_validado if telefone_validado else ""

        if is_cnpj:
            planilha.at[indice_linha, "SOCIOEmail"] = emails[0] if len(emails) > 0 else ""
        else:
            planilha.at[indice_linha, "Email"] = emails[0] if len(emails) > 0 else ""

        if "Documento" in primeiro_socio and "Nome" in primeiro_socio:
            planilha.at[indice_linha, "SocioDocumento"] = primeiro_socio["Documento"]
            planilha.at[indice_linha, "SocioNome"] = primeiro_socio["Nome"]
        else:
            planilha.at[indice_linha, "SocioDocumento"] = ""
            planilha.at[indice_linha, "SocioNome"] = ""

    # Replicação das linhas para os outros sócios (SOCIO2 e SOCIO3)
    for i, socio in enumerate(socios_unicos[1:], start=2):  # SOCIO2 e SOCIO3
        telefone_validado = validar_telefone(socio["Telefone"])
        if telefone_validado:
            nova_linha = planilha.loc[indice_linha].copy()
            nova_linha["Telefone"] = telefone_validado
            if is_cnpj:
                nova_linha[f"SOCIO{i}Email"] = emails[i - 1] if len(emails) >= i else ""
            nova_linha["SocioDocumento"] = socio.get("Documento", "")
            nova_linha["SocioNome"] = socio.get("Nome", "")

            planilha = pd.concat([planilha, pd.DataFrame([nova_linha])], ignore_index=True)

    return planilha

# Loop através dos arquivos no diretório
for arquivo in arquivos_no_diretorio:
    if arquivo.endswith('.xlsx'):
        planilha_diretorio = pd.read_excel(os.path.join(diretorio, arquivo))

        planilha_enriquecimento_cpf = pd.read_csv('cpf.txt_LAYOUT_Assertiva_V2_Whatsapp_PF.csv', sep=';', encoding='latin-1')
        planilha_enriquecimento_cnpj = pd.read_csv('cnpj.txt_LAYOUT_Assertiva_PJ_Socios_e_Relacionadas_PJ.csv', sep=';', encoding='latin-1')


        if "Telefone" not in planilha_diretorio.columns:
            planilha_diretorio["Telefone"] = ""  # Cria a coluna com valores vazios

        for indice_linha, cpfcnpj_celula in enumerate(planilha_diretorio["CPF/CNPJ"]):
            cpf_cnpj_formatado = str(cpfcnpj_celula).replace('-', '').replace('/', '').replace('.', '')

            email = ''
            socios_com_telefone = []  # Lista para armazenar apenas sócios que possuem telefone

            if not pd.isna(cpfcnpj_celula):
                # Se for CNPJ, buscar informações de sócios
                if len(cpf_cnpj_formatado) == 14:
                    socios = planilha_enriquecimento_cnpj[planilha_enriquecimento_cnpj["CNPJ"] == int(cpf_cnpj_formatado)]
                    if not socios.empty:
                        email = socios["Email1"].values[0] if len(socios["Email1"].values) > 0 else ""

                        # Coletar apenas sócios que têm telefone preenchido
                        for i in range(1, 4):  # SOCIO1, SOCIO2, SOCIO3
                            telefone = str(socios[f"SOCIO{i}Celular1"].values[0]).replace(".0", "").strip()
                            if telefone and telefone.lower() not in ('nan', 'nannan', ''):
                                socios_com_telefone.append({
                                    "Documento": str(socios[f"SOCIO{i}Documento"].values[0]).strip(),
                                    "Nome": str(socios[f"SOCIO{i}Nome"].values[0]).strip(),
                                    "Telefone": f"+55{telefone}"
                                })
                    # Passar True para indicar que é CNPJ
                    planilha_diretorio = replicar_linhas(planilha_diretorio, indice_linha, socios_com_telefone, [email], True)

                # Se for CPF, buscar dados na planilha de CPF
                elif len(cpf_cnpj_formatado) == 11:
                    pessoa_fisica = planilha_enriquecimento_cpf[planilha_enriquecimento_cpf["CPF"] == int(cpf_cnpj_formatado)]
                    if not pessoa_fisica.empty:
                        # Coletar todos os telefones (Celular1, Celular2 e Celular3)
                        telefones = [str(t).replace(".0", "").strip() for t in pessoa_fisica[["Celular1", "Celular2", "Celular3"]].values.flatten().tolist()]
                        email = pessoa_fisica["Email1"].values[0] if len(pessoa_fisica["Email1"].values) > 0 else ""

                        # Preencher o Email1 da pessoa física (CPF) apenas na coluna correta
                        planilha_diretorio.at[indice_linha, "Email"] = email if email else ""

                        # Adicionar os telefones à lista de sócios (sem documento e nome)
                        socios_com_telefone = [{"Telefone": f"+55{t}"} for t in telefones if str(t).lower() not in ('nan', 'nannan', '')]

                    # Passar False para indicar que é CPF
                    planilha_diretorio = replicar_linhas(planilha_diretorio, indice_linha, socios_com_telefone, [email], False)

        # Filtrar linhas com telefone preenchido
        planilha_diretorio = planilha_diretorio[planilha_diretorio['Telefone'].notna() & (planilha_diretorio['Telefone'] != '')]

        # Salvar a planilha atualizada
        planilha_diretorio.to_excel(os.path.join(diretorio, arquivo), sheet_name="Plan1", index=False)

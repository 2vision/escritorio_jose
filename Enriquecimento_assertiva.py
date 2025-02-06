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

def formatar_cpf(cpf):
    if pd.isna(cpf) or str(cpf).strip().lower() in ('nan', ''):
        return ""
    cpf = cpf.replace('.0', '')
    if len(cpf) == 11:
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    else:
        cpf = cpf.zfill(11)
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

def replicar_linhas(planilha, indice_linha, socios_com_telefone, email, is_cnpj):
    if not socios_com_telefone:
        return planilha

    primeiro_socio = socios_com_telefone[0]
    telefone_validado = validar_telefone(primeiro_socio["Telefone"])

    if "Telefone" not in planilha.columns:
        planilha["Telefone"] = ""

    planilha.at[indice_linha, "Telefone"] = telefone_validado if telefone_validado else ""

    # Usar apenas uma coluna para e-mail do sócio
    planilha.at[indice_linha, "SOCIOEmail"] = email if email else ""

    # Formatar documento corretamente apenas se for CPF (11 dígitos)
    documento = str(primeiro_socio.get("Documento", "")).strip()

    documento = formatar_cpf(documento)

    planilha.at[indice_linha, "SocioDocumento"] = documento
    planilha.at[indice_linha, "SocioNome"] = primeiro_socio.get("Nome", "")

    return planilha



# Loop pelos arquivos no diretório
for arquivo in arquivos_no_diretorio:
    if arquivo.endswith('.xlsx'):
        planilha_diretorio = pd.read_excel(os.path.join(diretorio, arquivo))
        planilha_enriquecimento_cpf = pd.read_csv('cpf.txt_LAYOUT_Assertiva_V2_Whatsapp_PF.csv', sep=';', encoding='latin-1')
        planilha_enriquecimento_cnpj = pd.read_csv('cnpj.txt_LAYOUT_Assertiva_PJ_Socios_e_Relacionadas_PJ.csv', sep=';', encoding='latin-1')

        if "Telefone" not in planilha_diretorio.columns:
            planilha_diretorio["Telefone"] = ""

        for indice_linha, cpfcnpj_celula in enumerate(planilha_diretorio["CPF/CNPJ"]):
            cpf_cnpj_formatado = str(cpfcnpj_celula).replace('-', '').replace('/', '').replace('.', '')

            email = ''
            socios_com_telefone = []

            if not pd.isna(cpfcnpj_celula):
                if len(cpf_cnpj_formatado) == 14:
                    socios = planilha_enriquecimento_cnpj[planilha_enriquecimento_cnpj["CNPJ"] == int(cpf_cnpj_formatado)]
                    if not socios.empty:
                        email = socios["Email1"].values[0] if len(socios["Email1"].values) > 0 else ""
                        for i in range(1, 4):
                            telefone = str(socios[f"SOCIO{i}Celular1"].values[0]).replace(".0", "").strip()
                            if telefone and telefone.lower() not in ('nan', 'nannan', ''):
                                socios_com_telefone.append({
                                    "Documento": str(socios[f"SOCIO{i}Documento"].values[0]).strip(),
                                    "Nome": str(socios[f"SOCIO{i}Nome"].values[0]).strip(),
                                    "Telefone": f"+55{telefone}"
                                })
                    planilha_diretorio = replicar_linhas(planilha_diretorio, indice_linha, socios_com_telefone, email, True)
                elif len(cpf_cnpj_formatado) <= 11:
                    cpf_cnpj_formatado = cpf_cnpj_formatado.zfill(11)  # Garante que tenha 11 dígitos
                    pessoa_fisica = planilha_enriquecimento_cpf[planilha_enriquecimento_cpf["CPF"] == int(cpf_cnpj_formatado)]
                    if not pessoa_fisica.empty:
                        telefones = [str(t).replace(".0", "").strip() for t in pessoa_fisica[["Celular1", "Celular2", "Celular3"]].values.flatten().tolist()]
                        email = pessoa_fisica["Email1"].values[0] if len(pessoa_fisica["Email1"].values) > 0 else ""
                        planilha_diretorio.at[indice_linha, "Email"] = email if email else ""
                        socios_com_telefone = [{"Telefone": f"+55{t}"} for t in telefones if str(t).lower() not in ('nan', 'nannan', '')]
                    planilha_diretorio = replicar_linhas(planilha_diretorio, indice_linha, socios_com_telefone, email, False)

        # Criar colunas CPF e CNPJ
        planilha_diretorio["CPF"] = planilha_diretorio["CPF/CNPJ"].apply(
            lambda x: x if len(str(x).replace(".", "").replace("-", "").replace("/", "")) == 11 else "")
        planilha_diretorio["CNPJ"] = planilha_diretorio["CPF/CNPJ"].apply(
            lambda x: x if len(str(x).replace(".", "").replace("-", "").replace("/", "")) == 14 else "")

        planilha_diretorio = planilha_diretorio[planilha_diretorio['Telefone'].notna() & (planilha_diretorio['Telefone'] != '')]
        planilha_diretorio.to_excel(os.path.join(diretorio, arquivo), sheet_name="Plan1", index=False)




from urllib.parse import quote
import pandas as pd
import os
from openpyxl.reader.excel import load_workbook

# Listar todos os arquivos no diretório
diretorio = 'Enriquecer'
arquivos_no_diretorio = os.listdir(diretorio)

# Função para validar e formatar o telefone
def validar_telefone(telefone):
    # Remove o prefixo '+'
    telefone = telefone.replace('+', '')

    # Verifica se o telefone tem pelo menos 12 caracteres (55 + DDD + número)
    if len(telefone) < 12:
        return ''

    # Extraindo o código de área (DDD) e o número
    codigo_pais = telefone[:2]
    ddd = telefone[2:4]
    numero = telefone[4:]

    # Ajustando o número de acordo com o DDD
    if int(ddd) <= 28:
        # Para DDD <= 28, o telefone deve ter 9 dígitos
        if len(numero) == 8:
            numero = '9' + numero
        elif len(numero) == 9 and numero[0] != '9':
            numero = '9' + numero[1:]
    else:
        # Para DDD >= 29, o telefone deve ter 8 dígitos
        if len(numero) == 9 and numero[0] == '9':
            numero = numero[1:]

    # Retorna o telefone formatado
    return f'+{codigo_pais}{ddd}{numero}'

# Função para replicar linhas
def replicar_linhas(planilha, indice_linha, telefones, email):
    for telefone in telefones:
        if telefone and telefone != 'nan':  # Certificar que o telefone é válido
            telefone_validado = validar_telefone(telefone)  # Validação do telefone
            if telefone_validado:  # Apenas replicar se o telefone for válido
                nova_linha = planilha.loc[indice_linha].copy()  # Copiar a linha atual
                nova_linha["Telefone1"] = telefone_validado  # Atribuir o telefone validado
                nova_linha["Email"] = email[0] if len(email) > 0 else ''  # Atribuir o email
                planilha = pd.concat([planilha, pd.DataFrame([nova_linha])], ignore_index=True)  # Adicionar a nova linha
    return planilha

# Loop através dos arquivos no diretório
for arquivo in arquivos_no_diretorio:
    # Verificar se o arquivo é uma planilha
    if arquivo.endswith('.xlsx'):
        # Carregar a planilha do diretório
        planilha_diretorio = pd.read_excel(os.path.join(diretorio, arquivo))

        # Carregar as planilhas de enriquecimento de CPF e CNPJ
        planilha_enriquecimento_cpf = pd.read_csv('cpf.txt_LAYOUT_Assertiva_V2_Whatsapp_PF.csv', sep=';', encoding='latin-1')
        planilha_enriquecimento_cnpj = pd.read_csv('cnpj.txt_LAYOUT_Assertiva_PJ_Socios_e_Relacionadas_PJ.csv', sep=';', encoding='latin-1')

        for indice_linha, cpfcnpj_celula in enumerate(planilha_diretorio["CPF/CNPJ"]):
            cpf_cnpj_formatado = str(cpfcnpj_celula).replace('-', '').replace('/', '').replace('.', '')

            telefones = []
            email = ''

            if not pd.isna(cpfcnpj_celula):
                # Consultar na planilha de enriquecimento CNPJ
                socios = planilha_enriquecimento_cnpj[planilha_enriquecimento_cnpj["CNPJ"] == int(cpf_cnpj_formatado)]
                if not socios.empty:
                    telefones = socios[["SOCIO1Celular1", "SOCIO2Celular1", "SOCIO3Celular1"]].values.flatten().tolist()
                    email = socios["Email1"].values

                # Se não encontrou no CNPJ, consultar na planilha de CPF
                if len(telefones) == 0:
                    pessoa_fisica = planilha_enriquecimento_cpf[planilha_enriquecimento_cpf["CPF"] == int(cpf_cnpj_formatado)]
                    if not pessoa_fisica.empty:
                        telefones = pessoa_fisica[["Celular1", "Celular2", "Celular3"]].values.flatten().tolist()
                        email = pessoa_fisica["Email1"].values

                # Filtrar os telefones válidos (remover 'nan' e valores vazios)
                telefones = [str(telefone).replace(".0", '') for telefone in telefones if telefone not in ('', 'nannan', 'nan')]
                telefones = [f'+55{telefone}' for telefone in telefones if telefone and telefone != 'nan']

                # Replicar as linhas para cada telefone após validação
                planilha_diretorio = replicar_linhas(planilha_diretorio, indice_linha, telefones, email)

        # Filtrar linhas que têm telefone preenchido
        planilha_diretorio = planilha_diretorio[
            planilha_diretorio['Telefone1'].notna() & (planilha_diretorio['Telefone1'] != '')]

        # Salvar a planilha atualizada
        planilha_diretorio.to_excel(os.path.join(diretorio, arquivo), sheet_name="Plan1", index=False)

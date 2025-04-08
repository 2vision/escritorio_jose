import pandas as pd
import os

diretorio = 'Enriquecer'
arquivos_no_diretorio = os.listdir(diretorio)


def formatar_cpf(cpf):
    cpf = str(cpf).replace('.0', '').strip()
    if pd.isna(cpf) or cpf.lower() in ('nan', ''):
        return ""
    cpf = cpf.zfill(11)
    if len(cpf) != 11 or not cpf.isdigit():
        return ""

    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


for arquivo in arquivos_no_diretorio:
    if arquivo.endswith('.xlsx'):
        caminho_arquivo = os.path.join(diretorio, arquivo)
        planilha_diretorio = pd.read_excel(caminho_arquivo)
        for coluna in ["NomeSocio", "TelefoneSocio", "DocumentoSocio"]:
            if coluna not in planilha_diretorio.columns:
                planilha_diretorio[coluna] = ""

        planilha_enriquecimento_cnpj = pd.read_csv('cnpj.txt_LAYOUT_Assertiva_PJ_Socios_e_Relacionadas_PJ.csv', sep=';',
                                                   encoding='latin-1')
        planilha_enriquecimento_cpf = pd.read_csv('cpf.txt_LAYOUT_Assertiva_V2_Whatsapp_PF.csv', sep=';',
                                                  encoding='latin-1')
        novas_linhas = []

        for indice_linha, linha in planilha_diretorio.iterrows():
            cpf_cnpj = str(linha["CPF/CNPJ"]).replace('-', '').replace('/', '').replace('.', '').strip()

            if len(cpf_cnpj) == 11:
                cpf_formatado = formatar_cpf(cpf_cnpj)
                telefones = planilha_enriquecimento_cpf[planilha_enriquecimento_cpf["CPF"] == int(cpf_cnpj)]

                if not telefones.empty:
                    telefone = telefones.iloc[0]["Celular1"]
                    if pd.notna(telefone) and str(telefone).strip():  # Garante que o telefone não seja vazio ou NaN
                        planilha_diretorio.at[indice_linha, "TelefoneSocio"] = f"+55{str(telefone).strip()}".replace('.0' , '')

            elif len(cpf_cnpj) == 14:
                cnpj_formatado = int(cpf_cnpj)
                socios = planilha_enriquecimento_cnpj[planilha_enriquecimento_cnpj["CNPJ"] == cnpj_formatado]

                if not socios.empty:
                    for i in range(1, 4):
                        nome_socio = socios.iloc[0].get(f"SOCIO{i}Nome", "")
                        telefone_socio = socios.iloc[0].get(f"SOCIO{i}Celular1", "")
                        documento_socio = str(socios.iloc[0].get(f"SOCIO{i}Documento", ""))

                        if pd.notna(telefone_socio) and str(telefone_socio).strip():
                            if nome_socio and not pd.isna(nome_socio) and str(nome_socio).strip().lower() not in (
                            'nan', ''):
                                if i == 1:
                                    planilha_diretorio.at[indice_linha, "NomeSocio"] = nome_socio
                                    planilha_diretorio.at[indice_linha, "TelefoneSocio"] = f"+55{str(telefone_socio).strip()}".replace('.0' , '')
                                    planilha_diretorio.at[indice_linha, "DocumentoSocio"] = formatar_cpf(
                                        documento_socio)
                                else:
                                    # Cria uma nova linha apenas se o telefone for válido
                                    nova_linha = linha.copy()
                                    nova_linha["NomeSocio"] = nome_socio
                                    nova_linha["TelefoneSocio"] = f"+55{str(telefone_socio).strip()}".replace('.0' , '')
                                    nova_linha["DocumentoSocio"] = formatar_cpf(documento_socio)
                                    novas_linhas.append(nova_linha)

        if novas_linhas:
            planilha_diretorio = pd.concat([planilha_diretorio, pd.DataFrame(novas_linhas)], ignore_index=True)

        planilha_diretorio.to_excel(caminho_arquivo, sheet_name="Plan1", index=False)

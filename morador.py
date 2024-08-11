import pandas as pd
import time 

start_time = time.time()

planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'
xls = pd.ExcelFile(planilha)

df_morador = pd.read_excel(xls, sheet_name='MORADOR')
df_deslocamento = pd.read_excel(xls, sheet_name='DESLOCAMENTO')
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')

df_morador = df_morador.merge(df_domicilio[['ID_DOMICILIO', 'ID_MENU']], 
                              on='ID_DOMICILIO', 
                              how='left')

deslocamento_count = df_deslocamento['ID_MORADOR'].value_counts()

df_morador['verificação deslocamentos cadastrados'] = df_morador.apply(
    lambda row: 'OK' if row['QUANTOS DESLOCAMENTOS VOCÊ FEZ ONTEM ?'] == deslocamento_count.get(row['ID_MORADOR'], 0) 
    else deslocamento_count.get(row['ID_MORADOR'], 0), 
    axis=1
)

def verifica_estudante(row):
    if row['ESTUDANTE?'] == 'SIM':
        required_columns = [
            'NOME DO LOCAL ONDE ESTUDA', 
            'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO ONDE ESTUDA ?', 
            'NOME DO LOGRADOURO ONDE ESTUDA ?', 
            'BAIRRO ONDE ESTUDA', 
            'MUNICÍPIO ONDE ESTUDA', 
            'QUAL O TIPO DE INSTITUIÇÃO ?'
        ]
        missing = [col for col in required_columns if pd.isna(row[col]) or row[col] == '']
        if len(missing) == len(required_columns):
            return 'INVÁLIDO'
        return 'OK' if not missing else '/'.join(missing)
    return '--'

df_morador['verificação estudante'] = df_morador.apply(verifica_estudante, axis=1)

def verifica_trabalho(row):
    if row['TRABALHO HOME OFFICE ?'] in ['SIM. 100% ONLINE, NÃO ME DESLOCO PARA O LOCAL DE TRABALHO', 'PARCIALMENTE. FAÇO HOME OFFICE, MAS ME DESLOCO PARA O LOCAL DE TRABALHO EM ALGUNS DIAS.']:
        return 'TRABALHO HOME OFFICE'
    if row['CONDIÇÃO DE ATIVIDADE PRINCIPAL'] in ['TRABALHO FIXO O DIA TODO', 'TRABALHO FIXO DE MEIO PERÍODO']:
        required_columns = [
            'NOME DO TRABALHO', 
            'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO DO TRABALHO', 
            'NOME DO LOGRADOURO(TRABALHO)', 
            'QUAL OUTRA RUA ?', 
            'BAIRRO(TRABALHO)', 
            'MUNICÍPIO(TRABALHO)'
        ]
        missing = [col for col in required_columns if pd.isna(row[col]) or row[col] == '']
        if len(missing) == len(required_columns):
            return 'INVÁLIDO'
        return 'OK' if not missing else '/'.join(missing)
    return '--'

df_morador['verificação trabalho'] = df_morador.apply(verifica_trabalho, axis=1)

df_morador['ENDEREÇO ESCOLA CONCETANADO'] = df_morador.apply(
    lambda row: f"{row['NOME DO LOCAL ONDE ESTUDA']} - {row['NOME DO LOGRADOURO ONDE ESTUDA ?']}, {row['BAIRRO ONDE ESTUDA']} - {row['MUNICÍPIO ONDE ESTUDA']} /RR"
    if row['ESTUDANTE?'] == 'SIM' else '--',
    axis=1
)


def verifica_trabalho(row):
    if row['CONDIÇÃO DE ATIVIDADE PRINCIPAL'] in ['TRABALHO FIXO O DIA TODO', 'TRABALHO FIXO DE MEIO PERÍODO', 'TRABALHO TEMPORÁRIO/BICO']:
        endereco_trabalho = f"{row['NOME DO TRABALHO']} - {row['NOME DO LOGRADOURO(TRABALHO)']}, {row['BAIRRO(TRABALHO)']} - {row['MUNICÍPIO(TRABALHO)']} /RR"
        return endereco_trabalho
    return '--'

def verifica_trabalho_secundario(row):
    if row['CONDIÇÃO DA ATIVIDADE SECUNDÁRIA'] in ['TRABALHO TEMPORÁRIO/BICO', 'TRABALHO FIXO O DIA TODO', 'TRABALHO FIXO DE MEIO PERÍODO']:
        endereco_trabalho_2 = f"{row['NOME DO LOCAL - TRABALHO (2)']} - {row['NOME DO LOGRADOURO - TRABALHO(2)']} {row['NÚMERO DO LOGRADOURO - TRABALHO(2)']}, {row['BAIRRO - TRABALHO(2)']} - {row['MUNICÍPIO - TRABALHO(2)']} /RR"
        return endereco_trabalho_2
    return '--'

df_morador['ENDEREÇO TRABALHO 2'] = df_morador.apply(verifica_trabalho_secundario, axis=1)

df_morador['ENDEREÇO TRABALHO 1'] = df_morador.apply(verifica_trabalho, axis=1)

with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_morador.to_excel(writer, sheet_name='MORADOR', index=False)

end_time = time.time()
execution_time = end_time - start_time 

print(execution_time)
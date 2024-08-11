import pandas as pd
import time 

start_time = time.time()

planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'
xls = pd.ExcelFile(planilha)

existing_sheets = pd.read_excel(planilha, sheet_name=None)

df_morador = pd.read_excel(planilha, sheet_name='MORADOR')
df_deslocamento = pd.read_excel(planilha, sheet_name='DESLOCAMENTO')

def format_origem(row):
    if 'SEU DOMICÍLIO' in row['ORIGEM']:
        return 'SEU DOMICÍLIO'
    elif 'TRABALHO 1' in row['ORIGEM']:
        morador_info = df_morador.loc[df_morador['ID_MORADOR'] == row['ID_MORADOR'], 'ENDEREÇO TRABALHO 1'].values
        return morador_info[0] if len(morador_info) > 0 else ''
    elif 'TRABALHO 2' in row['ORIGEM']:
        morador_info = df_morador.loc[df_morador['ID_MORADOR'] == row['ID_MORADOR'], 'ENDEREÇO TRABALHO 2'].values
        return morador_info[0] if len(morador_info) > 0 else ''
    elif 'OUTROS' in row['ORIGEM']:
        return f"{row['PONTO DE REFERÊNCIA']} - {row['ENDEREÇO DE ORIGEM']} {row['QUAL O NÚMERO DO LOCAL DE ORIGEM ?']}, {row['MUNICÍPIO DA ORIGEM']}, {row['QUAL O MUNICÍPIO DO LOCAL DE ORIGEM ?']}"
    elif 'ESCOLA / FACULDADE / CURSO' in row['ORIGEM']:
        morador_info = df_morador.loc[df_morador['ID_MORADOR'] == row['ID_MORADOR'], 'ENDEREÇO ESCOLA CONCETANADO'].values
        return morador_info[0] if len(morador_info) > 0 else ''
    else:
        return ''

df_deslocamento['ENDEREÇO_ORIGEM_BRUTO'] = df_deslocamento.apply(format_origem, axis=1)

def format_destino(row):
    if 'SEU DOMICÍLIO' in row['QUANTIDADE DE VIAGENS']:
        return 'SEU DOMICÍLIO'
    elif 'TRABALHO 1' in row['QUANTIDADE DE VIAGENS']:
        morador_info = df_morador.loc[df_morador['ID_MORADOR'] == row['ID_MORADOR'], 'ENDEREÇO TRABALHO 1'].values
        return morador_info[0] if len(morador_info) > 0 else ''
    elif 'TRABALHO 2' in row['QUANTIDADE DE VIAGENS']:
        morador_info = df_morador.loc[df_morador['ID_MORADOR'] == row['ID_MORADOR'], 'ENDEREÇO TRABALHO 2'].values
        return morador_info[0] if len(morador_info) > 0 else ''
    elif 'OUTROS' in row['QUANTIDADE DE VIAGENS']:
        return f"{row['PONTO DE REFERÊNCIA(DESTINO)']} - {row['QUAL O LOGRADOURO DO LOCAL DO DESTINO']} {row['QUAL O NÚMERO DO LOCAL ?']}, {row['QUAL O BAIRRO DO LOCAL(DESTINO)']}, {row['MUNICÍPIO DO LOCAL DE DESTINO']}"
    elif 'ESCOLA / FACULDADE / CURSO' in row['QUANTIDADE DE VIAGENS']:
        morador_info = df_morador.loc[df_morador['ID_MORADOR'] == row['ID_MORADOR'], 'ENDEREÇO ESCOLA CONCETANADO'].values
        return morador_info[0] if len(morador_info) > 0 else ''
    else:
        return ''

df_deslocamento['ENDEREÇO_DESTINO_BRUTO'] = df_deslocamento.apply(format_destino, axis=1)

existing_sheets['MORADOR'] = df_morador
existing_sheets['DESLOCAMENTO'] = df_deslocamento

with pd.ExcelWriter(planilha, engine='openpyxl') as writer:
    for sheet_name, df in existing_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

end_time = time.time()
execution_time = end_time - start_time   
print(execution_time)     

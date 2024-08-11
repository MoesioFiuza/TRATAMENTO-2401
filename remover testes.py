import pandas as pd
import time 

start_time = time.time()

planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'

df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')
df_morador = pd.read_excel(planilha, sheet_name='MORADOR')
df_deslocamento = pd.read_excel(planilha, sheet_name='DESLOCAMENTO')

ids_to_remove = df_menu[df_menu['IDENTIFICAÇÃO DO PESQUISADOR'].str.contains('teste', case=False, na=False)]['ID_MENU']

df_menu = df_menu[~df_menu['ID_MENU'].isin(ids_to_remove)]
df_domicilio = df_domicilio[~df_domicilio['ID_MENU'].isin(ids_to_remove)]
df_morador = df_morador[~df_morador['ID_MENU'].isin(ids_to_remove)]
df_deslocamento = df_deslocamento[~df_deslocamento['ID_MENU'].isin(ids_to_remove)]

for df in [df_menu, df_domicilio, df_morador, df_deslocamento]:
    if 'DATA' in df.columns:
        df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')

with pd.ExcelWriter(planilha, engine='openpyxl') as writer:
    df_menu.to_excel(writer, sheet_name='MENU', index=False)
    df_domicilio.to_excel(writer, sheet_name='DOMICÍLIO', index=False)
    df_morador.to_excel(writer, sheet_name='MORADOR', index=False)
    df_deslocamento.to_excel(writer, sheet_name='DESLOCAMENTO', index=False)

end_time = time.time()
execution_time = end_time - start_time    

print(execution_time)
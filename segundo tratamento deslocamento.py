import pandas as pd
import time 

start_time = time.time()


planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'
xls = pd.ExcelFile(planilha)

existing_sheets = pd.read_excel(planilha, sheet_name=None)

df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')
df_morador = pd.read_excel(planilha, sheet_name='MORADOR')
df_deslocamento = pd.read_excel(planilha, sheet_name='DESLOCAMENTO')

df_deslocamento['ENDEREÇO ORIGEM GEOCODIFICADO'] = ''
df_deslocamento['LATITUDE ORIGEM'] = ''
df_deslocamento['LONGITUDE ORIGEM'] = ''
df_deslocamento['ENDEREÇO DESTINO GEOCODIFICADO'] = ''
df_deslocamento['LATITUDE DESTINO'] = ''
df_deslocamento['LONGITUDE DESTINO'] = ''

df_deslocamento = df_deslocamento.merge(df_morador[['ID_MORADOR', 'ID_DOMICILIO']], on='ID_MORADOR', how='left')

df_deslocamento = df_deslocamento.merge(df_domicilio[['ID_DOMICILIO', 'ID_MENU']], on='ID_DOMICILIO', how='left')

df_domicilio = df_domicilio.merge(df_menu[['ID_MENU', 'DATA']], on='ID_MENU', how='left')
df_morador = df_morador.merge(df_menu[['ID_MENU', 'DATA']], on='ID_MENU', how='left')
df_deslocamento = df_deslocamento.merge(df_menu[['ID_MENU', 'DATA']], on='ID_MENU', how='left')


with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    for sheet_name in existing_sheets.keys():
        if sheet_name == 'MENU':
            df_menu.to_excel(writer, sheet_name=sheet_name, index=False)
        elif sheet_name == 'DOMICÍLIO':
            df_domicilio.to_excel(writer, sheet_name=sheet_name, index=False)
        elif sheet_name == 'MORADOR':
            df_morador.to_excel(writer, sheet_name=sheet_name, index=False)
        elif sheet_name == 'DESLOCAMENTO':
            df_deslocamento.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            existing_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

end_time = time.time()
execution_time = end_time - start_time             
print(execution_time)
import pandas as pd
import time

start_time = time.time()

planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'
zonas = r'/home/moesiosf/Área de trabalho/TRATAMENTO/EnderecosDomiciliar_2401_01.xlsx'
logins = r'/home/moesiosf/Área de trabalho/TRATAMENTO/0_2_Logins.csv'

# Ler apenas as colunas necessárias
df_zonas = pd.read_excel(zonas, sheet_name='Domicilios', usecols=['IDENTIFICADOR DO DOMICÍLIO', 'ZONA', 'ENDEREÇO FORMATADO'])
df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_logins = pd.read_csv(logins, usecols=['ID_LOGIN', 'Nome do App'])

df_menu = df_menu.merge(df_zonas[['IDENTIFICADOR DO DOMICÍLIO', 'ZONA']], how='left', on='IDENTIFICADOR DO DOMICÍLIO')

df_menu['NÚMERO DO DOMICÍLIO'] = df_menu['NÚMERO DO DOMICÍLIO VIZINHO'].combine_first(df_menu['NÚMERO DO DOMICÍLIO'])
df_menu['NÚMERO DO DOMICÍLIO'] = df_menu['NÚMERO DO DOMICÍLIO'].fillna(0).astype(int).astype(str)

# Criar mapa de endereços formatados
endereco_formatado_map = df_zonas.set_index('IDENTIFICADOR DO DOMICÍLIO')['ENDEREÇO FORMATADO'].to_dict()

# Função vetorizada para obter endereço concatenado
def get_endereco_concatenado(row):
    if row['DOMICÍLIO REGISTRADO NO APLICATIVO?'] == 'SIM':
        return endereco_formatado_map.get(row['IDENTIFICADOR DO DOMICÍLIO'], '')
    else:
        return (str(row['NOME DO LOGRADOURO']) + ", " +
                row['NÚMERO DO DOMICÍLIO'] + ", " + 
                str(row['BAIRRO DO DOMICÍLIO']) + " - " +
                "BOA VISTA / RR").upper()

df_menu['ENDEREÇO CONCATENADO'] = df_menu.apply(get_endereco_concatenado, axis=1)

# Merge com df_logins
df_menu = df_menu.merge(df_logins, how='left', on='ID_LOGIN')
df_menu.rename(columns={'Nome do App': 'Versão do APP'}, inplace=True)

# Escrever apenas a planilha 'MENU'
with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_menu.to_excel(writer, sheet_name='MENU', index=False)

end_time = time.time()
execution_time = end_time - start_time

print(execution_time)
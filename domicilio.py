import pandas as pd
import time

start_time = time.time()

planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'

df_morador = pd.read_excel(planilha, sheet_name='MORADOR', usecols=['ID_DOMICILIO'])
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')
df_menu = pd.read_excel(planilha, sheet_name='MENU', usecols=['ID_MENU', 'DATA'])

contagem_morador = df_morador['ID_DOMICILIO'].value_counts().reset_index()
contagem_morador.columns = ['ID_DOMICILIO', 'QTD_PESSOAS_MORADOR']

df_domicilio = df_domicilio.merge(contagem_morador, on='ID_DOMICILIO', how='left')

df_domicilio['VALIDAÇÃO'] = df_domicilio['QUANTAS PESSOAS MORAM NESSE DOMICÍLIO ?'] == df_domicilio['QTD_PESSOAS_MORADOR']
ajustes_realizados = df_domicilio['VALIDAÇÃO'] == False
df_domicilio.loc[ajustes_realizados, 'QUANTAS PESSOAS MORAM NESSE DOMICÍLIO ?'] = df_domicilio.loc[ajustes_realizados, 'QTD_PESSOAS_MORADOR']
df_domicilio['AJUSTADO'] = ajustes_realizados.replace({True: 'SIM', False: 'NÃO'})

colunas_para_verificar = [
    'GELADEIRA', 'FREEZER', 'LAVA-LOUÇAS', 'MÁQUINA DE LAVAR ROUPA',
    'MÁQUINA DE SECAR ROUPA', 'MICRO-ONDAS', 'COMPUTADOR/NOTEBOOK',
    'DVD/VIDEOGAME', 'BANHEIROS', 'QUANTOS EMPREGADOS DOMÉSTICOS QUE TRABALHAM PELO MENOS 5 DIAS POR SEMANA',
    'BICICLETAS', 'MOTOCICLETAS', 'AUTOMÓVEIS', 'TEM ÁGUA ENCANADA NESTE DOMICÍLIO ?'
]
df_domicilio['VERIFICAÇÃO DOS ITENS DE CONFORTO PREENCHIDOS'] = df_domicilio[colunas_para_verificar].isna().apply(
    lambda row: '/'.join(row.index[row]) if row.any() else 'OK', axis=1
)

df_domicilio = df_domicilio.merge(df_menu, on='ID_MENU', how='left')

with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    for sheet_name in pd.ExcelFile(planilha).sheet_names:
        if sheet_name != 'DOMICÍLIO':
            pd.read_excel(planilha, sheet_name=sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
    df_domicilio.to_excel(writer, sheet_name='DOMICÍLIO', index=False)

end_time = time.time()
execution_time = end_time - start_time

print(execution_time)
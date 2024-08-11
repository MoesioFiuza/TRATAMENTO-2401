import pandas as pd
import time 

start_time = time.time()


planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'
xls = pd.ExcelFile(planilha)

df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')
df_morador = pd.read_excel(xls, sheet_name='MORADOR')

df_domicilio.rename(columns={'DATA': 'DATA_DOMICILIO'}, inplace=True)

df_morador = df_morador.merge(df_domicilio[['ID_DOMICILIO', 'DATA_DOMICILIO']], 
                              on='ID_DOMICILIO', 
                              how='left')

colunas_ordenadas = ['DATA_DOMICILIO'] + [col for col in df_morador.columns if col != 'DATA_DOMICILIO']
df_morador = df_morador[colunas_ordenadas]

colunas_ordenadas = [
    'ID_MENU', 'ID_DOMICILIO', 'ID_MORADOR', 'PRIMEIRO NOME',
    'QUANTOS DESLOCAMENTOS VOCÊ FEZ ONTEM ?', 'verificação deslocamentos cadastrados',
    'SITUAÇÃO FAMILIAR', 'OUTROS(1)', 'GÊNERO', 'IDADE', 'QUAL GRAU DE INSTRUÇÃO ?',
    'ESTUDANTE?', 'AULAS ONLINE ?', 'SABE INFORMAR O NOME DO LOCAL ONDE ESTUDA ?',
    'NOME DO LOCAL ONDE ESTUDA', 'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO ONDE ESTUDA ?',
    'NOME DO LOGRADOURO ONDE ESTUDA ?', 'QUAL NOME DA OUTRA RUA ?', 'BAIRRO ONDE ESTUDA',
    'MUNICÍPIO ONDE ESTUDA', 'QUAL O TIPO DE INSTITUIÇÃO ?', 'CONDIÇÃO DE ATIVIDADE PRINCIPAL',
    'TRABALHO HOME OFFICE ?', 'SABE INFORMAR O NOME DO LOCAL DE TRABALHO ?', 'NOME DO TRABALHO',
    'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO DO TRABALHO', 'NOME DO LOGRADOURO(TRABALHO)',
    'QUAL OUTRA RUA ?', 'BAIRRO(TRABALHO)', 'MUNICÍPIO(TRABALHO)', 'CONDIÇÃO DA ATIVIDADE SECUNDÁRIA',
    'TRABALHO (2) HOME OFFICE ?', 'NOME DO TRABALHO(2)', 'NOME DO LOCAL - TRABALHO (2)',
    'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO DO TRABALHO(2)', 'NOME DO LOGRADOURO - TRABALHO(2)',
    'NÚMERO DO LOGRADOURO - TRABALHO(2)', 'SABE INFORMAR O BAIRRO - TRABALHO(2)', 'BAIRRO - TRABALHO(2)',
    'MUNICÍPIO - TRABALHO(2)', 'OUTRO MUNICÍPIO - TRABALHO(2)', 'RENDA', 'POSSUI LIMITAÇÕES DE MOBILIDADE',
    'VOCÊ SE ENVOLVEU EM ALGUM ACIDENTE DE TRÂNSITO NOS ÚLTIMOS 12 MESES ?', 'QUAL MODO DE TRANSPORTE VOCÊ ESTAVA UTILIZANDO NO MOMENTO DO ACIDENTE ?',
    'OUTRO (2)', 'QUAL MODO DE TRANSPORTE DO(S) OUTRO(S) ENVOLVIDO(S) NO ACIDENTE ?', 'OUTRO (3)',
    'TEVE ALGUMA LESÃO PERMANENTE APÓS O ACIDENTE ?', 'MUNICÍPIO DO ACIDENTE', 'POR QUE VOCÊ NÃO FEZ NENHUM DESLOCAMENTO ONTEM ?',
    'OUTROS ?', 'NACIONALIDADE', 'OUTRO MUNICÍPIO - ESTUDO', 'NÚMERO DO LOGRADOURO ONDE ESTUDA',
    'OUTRO MUNICÍPIO - TRABALHO', 'NÚMERO DO LOGRADOURO - TRABALHO', 'BAIRRO DO ACIDENTE',
    'NOME DO LOGRADOURO DO ACIDENTE', 'NÚMERO DO LOGRADOURO DO ACIDENTE', 'RAÇA', 'SABE INFORMAR O BAIRRO ?',
    'verificação estudante', 'verificação trabalho', 'ENDEREÇO ESCOLA CONCETANADO', 'ENDEREÇO TRABALHO 1',
    'ENDEREÇO TRABALHO 2'
]

colunas_ordenadas_2 = [
    'ID_MENU','ID_DOMICILIO','TIPO DE DOMICÍLIO','QUANTAS PESSOAS MORAM NESSE DOMICÍLIO ?','QUANTAS FAMÍLIAS MORAM NESSE DOMICÍLIO ?','VERIFICAÇÃO DOS ITENS DE CONFORTO PREENCHIDOS',
    'AJUSTADO','GELADEIRA','FREEZER','LAVA-LOUÇAS','MÁQUINA DE LAVAR ROUPA','MÁQUINA DE SECAR ROUPA','MICRO-ONDAS','COMPUTADOR/NOTEBOOK','DVD/VIDEOGAME','BANHEIROS','QUANTOS EMPREGADOS DOMÉSTICOS QUE TRABALHAM PELO MENOS 5 DIAS POR SEMANA',
    'BICICLETAS','MOTOCICLETAS','AUTOMÓVEIS','TEM ÁGUA ENCANADA NESTE DOMICÍLIO ?','QTD_PESSOAS_MORADOR','VALIDAÇÃO'
]

df_morador = df_morador[colunas_ordenadas]
df_domicilio = df_domicilio[colunas_ordenadas_2]

with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_morador.to_excel(writer, sheet_name='MORADOR', index=False)
    df_domicilio.to_excel(writer, sheet_name='DOMICÍLIO', index=False)

end_time = time.time()
execution_time = end_time - start_time

print(execution_time)
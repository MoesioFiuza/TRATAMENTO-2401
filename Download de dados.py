import pandas as pd
import time

start_time = time.time()

# Define file paths
menu = r'/home/moesiosf/Área de trabalho/TRATAMENTO/0_3_Menu.csv'
domicilio = r'/home/moesiosf/Área de trabalho/TRATAMENTO/1_Dados_do_Domicílio.csv'
morador = r'/home/moesiosf/Área de trabalho/TRATAMENTO/2_Morador.csv'
deslocamento = r'/home/moesiosf/Área de trabalho/TRATAMENTO/3_Deslocamento.csv'
logins = r'/home/moesiosf/Área de trabalho/TRATAMENTO/0_2_Logins.csv'

# Define new column names
novos_nomes_menu = {
    'P_0_1_1': 'DOMICÍLIO REGISTRADO NO APLICATIVO?',
    'P_0_1_2': 'NOME DO LOGRADOURO',
    'P_0_1_3': 'NÚMERO DO DOMICÍLIO',
    'P_0_1_4': 'BAIRRO DO DOMICÍLIO',
    'P_0_1': 'IDENTIFICAÇÃO DO PESQUISADOR',
    'P_0_2': 'IDENTIFICADOR DO DOMICÍLIO',
    'P_0_3': 'VERIFICADOR DOMICÍLIO',
    'Data': 'DATA',
    'Lat': 'LATITUDE',
    'Long': 'LONGITUDE',
    'Email': 'EMAIL',
    'NUMERO DA VISITA': 'NÚMERO DA VISITA',
    'P_0_4': 'TRECHO DA RUA DO DOMICÍLIO TEM PAVIMENTAÇÃO?',
    'P_0_7': 'PESQUISA EM DOMICÍLIO VIZINHO?',
    'P_0_8': 'NÚMERO DO DOMICÍLIO VIZINHO',
    'P_0_5': 'RESULTADO PRELIMINAR DA PESQUISA',
    'P_0_6': 'TELEFONE PARA CONTATO (DDD+NÚMERO)'
}

novos_nomes_domicilio = {
    'P_1_1': 'TIPO DE DOMICÍLIO',
    'P_1_2': 'QUANTAS PESSOAS MORAM NESSE DOMICÍLIO ?',
    'P_1_3': 'QUANTAS FAMÍLIAS MORAM NESSE DOMICÍLIO ?',
    'P_1_4': 'GELADEIRA',
    'P_1_5': 'FREEZER',
    'P_1_6': 'LAVA-LOUÇAS',
    'P_1_7': 'MÁQUINA DE LAVAR ROUPA',
    'P_1_8': 'MÁQUINA DE SECAR ROUPA',
    'P_1_9': 'MICRO-ONDAS',
    'P_1_10': 'COMPUTADOR/NOTEBOOK',
    'P_1_11': 'DVD/VIDEOGAME',
    'P_1_12': 'BANHEIROS',
    'P_1_13': 'QUANTOS EMPREGADOS DOMÉSTICOS QUE TRABALHAM PELO MENOS 5 DIAS POR SEMANA',
    'P_1_14': 'BICICLETAS',
    'P_1_15': 'MOTOCICLETAS',
    'P_1_16': 'AUTOMÓVEIS',
    'P_1_17': 'TEM ÁGUA ENCANADA NESTE DOMICÍLIO ?'
}

novos_nomes_morador = {
    'P_2_1': 'PRIMEIRO NOME',
    'P_2_2': 'SITUAÇÃO FAMILIAR',
    'P_2_3': 'OUTROS(1)',
    'P_2_4': 'GÊNERO',
    'P_2_5': 'IDADE',
    'P_2_48': 'RAÇA',
    'P_2_5_1': 'NACIONALIDADE',
    'P_2_6': 'QUAL GRAU DE INSTRUÇÃO ?',
    'P_2_7': 'ESTUDANTE?',
    'P_2_8': 'AULAS ONLINE ?',
    'P_2_9': 'SABE INFORMAR O NOME DO LOCAL ONDE ESTUDA ?',
    'P_2_10': 'NOME DO LOCAL ONDE ESTUDA',
    'P_2_11': 'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO ONDE ESTUDA ?',
    'P_2_12': 'NOME DO LOGRADOURO ONDE ESTUDA ?',
    'P_2_12_1': 'NÚMERO DO LOGRADOURO ONDE ESTUDA',
    'P_2_13': 'QUAL NOME DA OUTRA RUA ?',
    'P_2_14_1': 'SABE INFORMAR O BAIRRO ?',
    'P_2_14': 'BAIRRO ONDE ESTUDA',
    'P_2_15': 'MUNICÍPIO ONDE ESTUDA',
    'P_2_15_1': 'OUTRO MUNICÍPIO - ESTUDO',
    'P_2_16': 'QUAL O TIPO DE INSTITUIÇÃO ?',
    'P_2_17': 'CONDIÇÃO DE ATIVIDADE PRINCIPAL',
    'P_2_18': 'TRABALHO HOME OFFICE ?',
    'P_2_19': 'SABE INFORMAR O NOME DO LOCAL DE TRABALHO ?',
    'P_2_20': 'NOME DO TRABALHO',
    'P_2_21': 'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO DO TRABALHO',
    'P_2_22_1': 'NÚMERO DO LOGRADOURO - TRABALHO',
    'P_2_22': 'NOME DO LOGRADOURO(TRABALHO)',
    'P_2_23': 'QUAL OUTRA RUA ?',
    'P_2_24': 'BAIRRO(TRABALHO)',
    'P_2_25': 'MUNICÍPIO(TRABALHO)',
    'P_2_25_1': 'OUTRO MUNICÍPIO - TRABALHO',
    'P_2_26': 'CONDIÇÃO DA ATIVIDADE SECUNDÁRIA',
    'P_2_27': 'TRABALHO (2) HOME OFFICE ?',
    'P_2_28': 'NOME DO TRABALHO(2)',
    'P_2_29': 'NOME DO LOCAL - TRABALHO (2)',
    'P_2_30': 'SABE INFORMAR O NOME E NÚMERO DO LOGRADOURO DO TRABALHO(2)',
    'P_2_31': 'NOME DO LOGRADOURO - TRABALHO(2)',
    'P_2_31_1': 'NÚMERO DO LOGRADOURO - TRABALHO(2)',
    'P_2_32': 'SABE INFORMAR O BAIRRO - TRABALHO(2)',
    'P_2_33': 'BAIRRO - TRABALHO(2)',
    'P_2_34': 'MUNICÍPIO - TRABALHO(2)',
    'P_2_34_1': 'OUTRO MUNICÍPIO - TRABALHO(2)',
    'P_2_35': 'RENDA',
    'P_2_36': 'POSSUI LIMITAÇÕES DE MOBILIDADE',
    'P_2_37': 'VOCÊ SE ENVOLVEU EM ALGUM ACIDENTE DE TRÂNSITO NOS ÚLTIMOS 12 MESES ?',
    'P_2_38': 'QUAL MODO DE TRANSPORTE VOCÊ ESTAVA UTILIZANDO NO MOMENTO DO ACIDENTE ?',
    'P_2_39': 'OUTRO (2)',
    'P_2_40': 'QUAL MODO DE TRANSPORTE DO(S) OUTRO(S) ENVOLVIDO(S) NO ACIDENTE ?',
    'P_2_41': 'OUTRO (3)',
    'P_2_42': 'TEVE ALGUMA LESÃO PERMANENTE APÓS O ACIDENTE ?',
    'P_2_43': 'MUNICÍPIO DO ACIDENTE',
    'P_2_44': 'BAIRRO DO ACIDENTE',
    'P_2_45': 'NOME DO LOGRADOURO DO ACIDENTE',
    'P_2_46': 'NÚMERO DO LOGRADOURO DO ACIDENTE',
    'P_3_1': 'QUANTOS DESLOCAMENTOS VOCÊ FEZ ONTEM ?',
    'P_3_38': 'POR QUE VOCÊ NÃO FEZ NENHUM DESLOCAMENTO ONTEM ?',
    'P_3_39': 'OUTROS ?'
}

novos_nomes_deslocamento = {
    'P_3_1': 'NOME DO LOGRADOURO',
    'P_3_2': 'ORIGEM',
    'P_3_3': 'INFORMOU PONTO DE REFERÊNCIA',
    'P_3_4': 'PONTO DE REFERÊNCIA',
    'P_3_5': 'SABE INFORMAR O LOCAL DE ORIGEM',
    'P_3_6': 'ENDEREÇO DE ORIGEM',
    'P_3_6_1': 'QUAL O NÚMERO DO LOCAL DE ORIGEM ?',
    'P_3_7': 'BAIRRO DA ORIGEM',
    'P_3_8': 'MUNICÍPIO DA ORIGEM',
    'P_3_8_1': 'QUAL O MUNICÍPIO DO LOCAL DE ORIGEM ?',
    'P_3_9': 'OUTRO MOTIVO(ORIGEM)',
    'P_3_10': 'MOTIVO (ORIGEM)',
    'P_3_11': 'DESTINO',
    'P_3_12': 'QUANTIDADE DE VIAGENS',
    'P_3_13': 'INFORMOU PONTO DE REFERÊNCIA(DESTINO)',
    'P_3_14': 'PONTO DE REFERÊNCIA(DESTINO)',
    'P_3_15': 'ENDEREÇO DO DESTINO',
    'P_3_16': 'QUAL O LOGRADOURO DO LOCAL DO DESTINO',
    'P_3_16_1': 'QUAL O NÚMERO DO LOCAL ?',
    'P_3_17': 'SABE INFOMAR O BAIRRO DO LOCAL(DESTINO)',
    'P_3_17_1': 'QUAL O BAIRRO DO LOCAL(DESTINO)',
    'P_3_17_2': 'MUNICÍPIO DO LOCAL DE DESTINO',
    'P_3_18': 'OUTRO MOTIVO(DESTINO)',
    'P_3_19': 'MOTIVO(DESTINO)',
    'P_3_20': 'QUE HORAS VOCÊ CHEGOU NO DESTINO ?',
    'P_3_21': 'NESSE DESLOCAMENTO, QUAL O PRINCIPAL MODO DE TRANSPORTE UTILIZADO? (AQUELE UTILIZADO POR MAIOR TEMPO/ MAIOR DISTÂNCIA)',
    'P_3_21_1': 'OUTROS',
    'P_3_22': 'NESSE DESLOCAMENTO, HOUVE OUTRO MODO DE TRANSPORTE?',
    'P_3_23': 'OUTROS',
    'P_3_24': 'QUANTOS MINUTOS ANDOU DO LOCAL DE ORIGEM AO PONTO DE ÔNIBUS ?',
    'P_3_25': 'QUANTOS MINUTOS ANDOU DO PONTO DE ÔNIBUS ATÉ LOCAL DE DESTINO?',
    'P_3_26': 'QUANTOS MINUTOS EM MÉDIA FICA ESPERANDO NO PONTO DE ÔNIBUS ?',
    'P_3_27': 'FORMA DE PAGAMENTO',
    'P_3_28': 'ONDE VOCÊ ESTACIONOU NO LOCAL DE DESTINO ?',
    'P_3_31': 'COMENTÁRIOS',
}

# Read CSV files and rename columns
menu_df = pd.read_csv(menu).rename(columns=novos_nomes_menu)
domicilio_df = pd.read_csv(domicilio).rename(columns=novos_nomes_domicilio)
morador_df = pd.read_csv(morador).rename(columns=novos_nomes_morador)
deslocamento_df = pd.read_csv(deslocamento).rename(columns=novos_nomes_deslocamento)

# Write to Excel
caminho_saida = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'
with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
    menu_df.to_excel(writer, sheet_name='MENU', index=False)
    domicilio_df.to_excel(writer, sheet_name='DOMICÍLIO', index=False)
    morador_df.to_excel(writer, sheet_name='MORADOR', index=False)
    deslocamento_df.to_excel(writer, sheet_name='DESLOCAMENTO', index=False)

end_time = time.time()
execution_time = end_time - start_time

print(execution_time)
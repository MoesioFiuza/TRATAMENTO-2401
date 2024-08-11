import pandas as pd

# Caminho do arquivo
planilha = r'C:\Users\moesios\Desktop\BD2401\BANCO DE DADOS VALIDADO (1).xlsx'

# Carregar os dados das abas 'MENU' e 'DESLOCAMENTO'
df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_deslocamento = pd.read_excel(planilha, sheet_name='DESLOCAMENTO')

# Função para verificar se todas as entradas de ID_MENU estão com ENDEREÇO GEOCODIFICADO e ENDEREÇO DESTINO GEOCODIFICADO preenchidos
def verificar_preenchimento(id_menu):
    deslocamentos = df_deslocamento[df_deslocamento['ID_MENU'] == id_menu]
    if deslocamentos[['ENDEREÇO GEOCODIFICADO', 'ENDEREÇO DESTINO GEOCODIFICADO']].notnull().all().all():
        return 1
    else:
        return 0

# Aplicar a função para cada ID_MENU na aba MENU
df_menu['Verificação'] = df_menu['ID_MENU'].apply(verificar_preenchimento)

# Caminho para salvar a nova planilha
saida = r'C:\Users\moesios\Desktop\BD2401\BANCO_DE_DADOS_VALIDADO_SAIDA.xlsx'

# Salvar a nova planilha com a saída
with pd.ExcelWriter(saida) as writer:
    df_menu.to_excel(writer, sheet_name='MENU', index=False)
    df_deslocamento.to_excel(writer, sheet_name='DESLOCAMENTO', index=False)

print(f'Nova planilha salva em: {saida}')



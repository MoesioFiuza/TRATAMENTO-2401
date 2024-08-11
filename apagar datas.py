import pandas as pd
import tkinter as tk
from tkinter import ttk

planilha = r'/home/moesiosf/Área de trabalho/TRATAMENTO/BD.xlsx'

# Carregar as planilhas
df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')
df_morador = pd.read_excel(planilha, sheet_name='MORADOR')
df_deslocamento = pd.read_excel(planilha, sheet_name='DESLOCAMENTO')

# Obter todas as datas únicas e ordenar
all_dates = pd.concat([df_menu['DATA'], df_domicilio['DATA'], df_morador['DATA'], df_deslocamento['DATA']]).dropna().unique()
all_dates = pd.to_datetime(all_dates, errors='coerce')  # Converter para datetime e tratar erros
all_dates = sorted(set(all_dates))  # Remover NaT e ordenar

# Função para aplicar o filtro
def apply_filter():
    selected_dates = [date for date, var in checkboxes.items() if var.get()]

    # Garantir que selected_dates não esteja vazio
    if not selected_dates:
        status_label.config(text="Nenhuma data selecionada.")
        return
    
    global df_menu, df_domicilio, df_morador, df_deslocamento
    
    # Manter todas as datas e apenas remover as linhas que não estão nas datas selecionadas
    df_menu_filtrado = df_menu[df_menu['DATA'].isin(selected_dates)]
    df_domicilio_filtrado = df_domicilio[df_domicilio['DATA'].isin(selected_dates)]
    df_morador_filtrado = df_morador[df_morador['DATA'].isin(selected_dates)]
    df_deslocamento_filtrado = df_deslocamento[df_deslocamento['DATA'].isin(selected_dates)]
    
    with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_menu_filtrado.to_excel(writer, sheet_name='MENU', index=False)
        df_domicilio_filtrado.to_excel(writer, sheet_name='DOMICÍLIO', index=False)
        df_morador_filtrado.to_excel(writer, sheet_name='MORADOR', index=False)
        df_deslocamento_filtrado.to_excel(writer, sheet_name='DESLOCAMENTO', index=False)
    
    status_label.config(text="Dados filtrados e salvos com sucesso.")

# Configuração da interface
root = tk.Tk()
root.title("Filtro de Dados")

# Criar checkboxes
checkboxes = {}
for date in all_dates:
    if pd.isna(date):  # Verifica se a data é NaT
        continue
    var = tk.BooleanVar(value=True)
    checkbox = tk.Checkbutton(root, text=date.strftime('%Y-%m-%d'), variable=var)
    checkbox.pack(anchor='w')
    checkboxes[date] = var

# Botão para aplicar o filtro
apply_button = tk.Button(root, text="Aplicar Filtro", command=apply_filter)
apply_button.pack(pady=10)

# Label de status
status_label = tk.Label(root, text="")
status_label.pack()

# Executar a interface
root.mainloop()
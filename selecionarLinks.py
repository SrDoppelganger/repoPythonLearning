import pandas as pd
import openpyxl

input_file = "selecao.xlsx"
df = pd.read_excel(input_file, sheet_name="Abreu e Lima",header=1)

#arrays de colunas
check_cols: list[int] = [9,12,15,18,21,24]
link_cols: list[int] = [7,10,13,16,19,22]

checkboxes = df.columns[check_cols].tolist()
links = df.columns[link_cols].tolist()

#check_values = df[df.columns[check_cols]]
#link_values = df[df.columns[link_cols]]

links_aceitos: list[str] = []

for i in range(6):
    check_values = df[df.columns[check_cols[i]]]
    link_values = df[df.columns[link_cols[i]]]

    for index,check in enumerate(check_values):
        if check:
            links_aceitos.append(link_values[index])



# Criação da planilha
wb = openpyxl.Workbook()
wb.create_sheet('links_aceitos')
wb.remove(wb['Sheet'])
wb['links_aceitos'].append(['Links'])
ws = wb.active
index_col = 1 #A -> 'links_aceitos'

for index,value in enumerate(links_aceitos, start=2):
    ws.cell(row=index, column=index_col, value=value)

wb.save("Links aceitos.xlsx")



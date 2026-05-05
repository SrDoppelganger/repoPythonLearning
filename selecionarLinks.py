import pandas as pd
import openpyxl

input_file = "selecao.xlsx"
df = pd.read_excel(input_file, sheet_name="Abreu e Lima",header=1)

#arrays de colunas
link_cols: list[int] = [7,10,13,16,19,22]
check_cols: list[int] = [9,12,15,18,21,24]

escolas_aceitas: list[str] = []
topicos_aceitos: list[str] = []
subtopicos_aceitos: list[str] = []
links_aceitos: list[str] = []

for i in range(6):
    escola_values = df[df.columns[1]]
    topico_values = df[df.columns[5]]
    subtopico_values = df[df.columns[6]]

    link_values = df[df.columns[link_cols[i]]]
    check_values = df[df.columns[check_cols[i]]]

    for index,check in enumerate(check_values):
        if check:
            escolas_aceitas.append(escola_values[index])
            topicos_aceitos.append(topico_values[index])
            subtopicos_aceitos.append(subtopico_values[index])
            links_aceitos.append(link_values[index])



# Criação da planilha
wb = openpyxl.Workbook()
wb.create_sheet('links_aceitos')
wb.remove(wb['Sheet'])
wb['links_aceitos'].append(['Unidade_Administrativa','Tópico','Subtópico','Links'])
ws = wb.active

index_col = 1 #A -> 'links_aceitos'

for row in zip(escolas_aceitas, topicos_aceitos, subtopicos_aceitos, links_aceitos):
    ws.append(row)

wb.save("Links aceitos.xlsx")



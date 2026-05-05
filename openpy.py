import openpyxl


input_file: str = "selecao.xlsx"
# Workbook = arquivo excel
wb = openpyxl.open(input_file)


checkboxes: list[str] = ['b']
#pegando uma linha de checkbox
municipio = wb['Abreu e Lima']
check_col = municipio[checkboxes[0]]
for cell in check_col:
    print(cell.value)

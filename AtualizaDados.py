from openpyxl import load_workbook

def convert_to_array(data_tuple):
    return [cell.value for cell in data_tuple]

spreadsheet1 = input("Nome da planilha que receberá os dados: ")
spreadsheet2 = input("Nome da planilha que contém os dados: ")
colkey = input("Coluna chave de transferência: ")

filename1 = spreadsheet1 + '.xlsx'
filename2 = spreadsheet2 + '.xlsx'

wb1 = load_workbook(filename=filename1)
wb2 = load_workbook(filename=filename2)

ws1 = wb1.active
ws2 = wb2.active

headers1 = convert_to_array(ws1['1'])
headers2 = convert_to_array(ws2['1'])

uniques = list(set(headers2) - set(headers1))

keys = convert_to_array(ws1[colkey])

for element in uniques:
    col_num = headers2.index(element) + 1
    col_data = ws2[col_num]
    for row in range(1, ws2.max_row + 1):
        key = ws2.cell(row=row,column=1).value
        if key in keys:
            rowidx = keys.index(key) + 1
            value = ws2.cell(row=row,column=col_num).value
            ws1.cell(row=rowidx,column=col_num).value = value

wb1.save(filename1)
wb2.save(filename2)
import openpyxl
from collections import OrderedDict
import json

workbook = openpyxl.load_workbook("ExportaCartillaProductos.xlsx")
sheet = workbook.active
data_list = []

for row in sheet.iter_rows():
    data = OrderedDict()
    data['docKey'] = int(row[0].value)
    data['name'] = row[1].value
    data['price'] = int(row[2].value)
    data_list.append(data)

with open("ToDB/data.json", "w", encoding="utf-8") as writeJsonfile:
    json.dump(data_list, writeJsonfile, indent=4, ensure_ascii=False, default=str)
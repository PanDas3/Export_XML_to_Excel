import xml.etree.ElementTree as ET
import pandas as pd
from xlsxwriter import Workbook

env = 'QA'
file = '*.xlsx'
xml = '*.xml'

tree = ET.parse(xml)
root = tree.getroot()

table = []
component = 'Globalny'
component2 = 'n\d'
for gp in root.findall('GlobalParams'):
    for params in gp:
        name = params.get('Symbol')
        value = params.find('TValue').text
        table.append([name, component, component2, value])
        
component = 'Składnikowy'
for gp in root.findall('ComponentialParams'):
    for params in gp:
        name = params.get('Symbol')
        component2 = params.get('Component')
        value = params.find('TValue').text
        table.append([name, component, component2, value])
        
component = 'Instancyjny'
for gp in root.findall('InstanceParams'):
    for params in gp:
        name = params.get('Symbol')
        machine = params.get('Machine')
        instance = params.get('Instance')
        component2 = machine + '\\' + instance
        value = params.find('TValue').text
        table.append([name, component, component2, value])
       
table = pd.DataFrame(table)
table.columns = 'Nazwa parametru', 'Typ', 'Komponent', 'Wartosc'
print(table)
 
writer = pd.ExcelWriter(file, engine='xlsxwriter')
table.to_excel(writer, sheet_name=env)

workbook = writer.book
worksheet = writer.sheets[env]
 
format1 = workbook.add_format()
format2 = workbook.add_format()
format3 = workbook.add_format()

format1.set_text_wrap()
format1.set_align('vcenter')
format2.set_align('vcenter')
format3.set_align('center')
format3.set_align('vcenter')

worksheet.set_column('E:E', 106.43, format1) #szerokosc kolumny z excel
worksheet.set_column('B:B', 74,29, format2)
worksheet.set_column('C:C', 13.57, format3)
worksheet.set_column('D:D', 50, format3)

writer.save()
import xml.etree.ElementTree as ET
import pandas as pd

output = 'D:\\Skrypty\\Parametry\\Parametry.xlsx'
envs = ["DEV", "QA", "PROD", "REF", "Diffrences"]
xmls = ['D:\\Skrypty\\Parametry\\DEV.xml', 'D:\\Skrypty\\Parametry\\QA.xml', 'D:\\Skrypty\\Parametry\\PROD.xml', 'D:\\Skrypty\\Parametry\\REF.xml']
df_env = []

for xml in xmls:

    tree = ET.parse(xml)
    root = tree.getroot()
    
    component = 'Global'
    component2 = 'n\d'
    table = []

    for gp in root.findall('GlobalParams'):
        for params in gp:
            name = params.get('Symbol')
            value = params.find('TValue').text
            table.append([name, component, component2, value])
            
    component = 'Componential'
    for gp in root.findall('ComponentialParams'):
        for params in gp:
            name = params.get('Symbol')
            component2 = params.get('Component')
            value = params.find('TValue').text
            table.append([name, component, component2, value])
            
    component = 'Instance'
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
    table = table.sort_values(['Nazwa parametru', 'Komponent', 'Wartosc'], ascending=[True, True, False])
    df_env.append(table)

differences = {"Nazwa parametru":[]}
for df in df_env:
    values = df["Nazwa parametru"].values
    for value in values:
        differences["Nazwa parametru"].append(value)

differences = pd.DataFrame(differences)
differences = differences.drop_duplicates(subset=['Nazwa parametru'])

counts = len(differences['Nazwa parametru'])
differences = differences.to_dict('list')

differences_tmp = {"Czy zgodne":[], "REF":[], "PROD":[], "QA":[], "DEV":[]}
for lp in range(2, counts +2, 1):
    differences_tmp['Czy zgodne'].append(f'=IFERROR(IF(AND(D{lp}=E{lp},D{lp}=F{lp},D{lp}=G{lp}),"Tak","Nie"),"Nie")')
    differences_tmp['PROD'].append(f'=VLOOKUP(B{lp},PROD!B:E,4,False')
    differences_tmp['REF'].append(f'=VLOOKUP(B{lp},REF!B:E,4,False')
    differences_tmp['QA'].append(f'=VLOOKUP(B{lp},QA!B:E,4,False')
    differences_tmp['DEV'].append(f'=VLOOKUP(B{lp},DEV!B:E,4,False')

differences.update(differences_tmp)
differences = pd.DataFrame(differences)
df_env.append(differences)

writer = pd.ExcelWriter(output, engine='xlsxwriter')
count_list = len(df_env)
count_list_iteration = 1

for env in envs:
    for df in df_env:
        print(df)
        df.to_excel(writer, sheet_name=env)

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

        if(count_list == count_list_iteration):
            worksheet.set_column('D:H', 50.0, format1)
            worksheet.set_column('B:B', 74.29, format2)
            worksheet.set_column('C:C', 13.57, format3)

        else:
            worksheet.set_column('E:E', 106.43, format1) #szerokosc kolumny z excel
            worksheet.set_column('B:B', 74,29, format2)
            worksheet.set_column('C:C', 13.57, format3)
            worksheet.set_column('D:D', 50, format3)

        count_list_iteration += 1
        break
    df_env.remove(df_env[0])

writer.save()
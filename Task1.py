import pandas as pd
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import openpyxl

# Открываем файл Excel
workbook = openpyxl.load_workbook('test_input.xlsx')

# Выбираем активный лист
sheet = workbook.active

# Чтение значения из ячейки и вычисление результата
value = sheet['B3'].value
result = sheet['B2'].value
date = sheet['B1'].value

# Соединяем значения вместе
result_formula = f"SABR0000001{date.strftime('%d%m%Y')}{result}"
# Чтение данных из файла Excel
df = pd.read_excel('test_input.xlsx', skiprows=4)

# Создание корневого элемента
certdata = ET.Element('CERTDATA')

# Добавление элемента FILENAME
filename = ET.SubElement(certdata, 'FILENAME')
filename.text = result_formula

# Создание элемента ENVELOPE
envelope = ET.SubElement(certdata, 'ENVELOPE')

# Создание элемента ECERT для каждой строки данных
for _, row in df.iterrows():
    ecert = ET.SubElement(envelope, 'ECERT')
    ET.SubElement(ecert, 'CERTNO').text = str(row['Ref no'])
    ET.SubElement(ecert, 'CERTDATE').text = str(row['Issuance Date']).split()[0]
    ET.SubElement(ecert, 'STATUS').text = str(row['Status'])
    ET.SubElement(ecert, 'IEC').text = str(row['IE Code'])
    ET.SubElement(ecert, 'EXPNAME').text = str(row['Client'])
    ET.SubElement(ecert, 'BILLID').text = str(row['Bill Ref no'])
    ET.SubElement(ecert, 'SDATE').text = str(row['SB Date']).split()[0]
    ET.SubElement(ecert, 'SCC').text = str(row['SB Currency'])
    ET.SubElement(ecert, 'SVALUE').text = str(row['SB Amount'])

# Создание и сохранение XML-файла
xml_str = minidom.parseString(ET.tostring(certdata)).toprettyxml(indent="\t", encoding='UTF-8')

# Запись в файл
with open("output_xml.xml", "w", encoding='UTF-8') as f:
    f.write(xml_str.decode())
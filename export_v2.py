from openpyxl import load_workbook
import xml.etree.ElementTree as ET
from datetime import datetime
import json,os,pytz, uuid,zipfile
import tkinter as tk
import tkinter.messagebox as messagebox
cashplan_namespace = "http://mil.ru/cashPlan"
ET.register_namespace('',cashplan_namespace)

def load_config(filename):
    with open(filename, "r", encoding='utf-8') as configfile:
        config_json = configfile.read()
    config = json.loads(config_json)
    return config

def current_datetime(value=None):
    tz = pytz.timezone("Europe/Moscow")
    if value is None:
     now_utc = datetime.now(pytz.UTC)
     now_local = now_utc.astimezone(tz).isoformat()
    else:
     now_local= tz.localize(datetime.strptime(str(value), "%d.%m.%Y")).replace(hour=0, minute=0, second=0).isoformat(timespec='seconds')
      
    return now_local

def load_workbook_data(filename):
    wb = load_workbook(filename=filename, read_only=True)
    return wb

def prettify_xml(element, level=0, indent='   '):
    if element:
        if not element.text or not element.text.strip():
            element.text = '\n' + (level+1) * indent
        if not element.tail or not element.tail.strip():
            element.tail = '\n' + level * indent
        for element in element:
            prettify_xml(element, level+1, indent)
    else:
        if level and (not element.tail or not element.tail.strip()):
            element.tail = '\n' + level * indent

def xml_to_pretty_string(root):
    prettify_xml(root)
    return ET.tostring(root, encoding='utf-8',xml_declaration=True,method='xml')

def add_to_zip(filepath,INN,UID):
    zip_name = os.path.dirname(filepath) + f"{INN}_{datetime.now().strftime('%Y%m%d')}_{UID}.zip"
    print(zip_name)
    with zipfile.ZipFile(zip_name,'w') as zipf:
       zipf.write(filepath)
       return filepath 
    
def find_last_row_with_prefix(sheet, start_row, prefix):
    last_row = None
    for row in range(start_row, sheet.max_row + 1):
        cell_value = sheet.cell(row=row, column=1).value
        # Предполагается, что значение является строкой и начинается с указанного префикса
        if str(cell_value).startswith(prefix):
            last_row = row
    return last_row

def find_first_row_with_prefix(sheet, start_row, prefix):
    for row in range(start_row, sheet.max_row + 1):
        cell_value = sheet.cell(row=row, column=1).value
        # Проверяем, что значение является строкой и начинается с указанного префикса
        if isinstance(cell_value, str) and cell_value.startswith(prefix):
            return row
    return None
def create_xml(sheet1,sheet2, config, now_local):
    thisUID = str(uuid.uuid4())

    root =ET.Element("Message")
    root.set('CreateDate',str(now_local))
    root.set('UID',str(thisUID))
    root.set('PreviousUID',str(config["PreviousUID"]))
    config["PreviousUID"] = thisUID
    Organization =ET.SubElement(root, "Organization")
    Organization.set('INN',str(config["INN"]))
    Organization.set("KPP",config["KPP"])
    Organization.set('Name',config["NAME"])
    Organization.set("GOZUID",str(sheet1[config["start_indexes"]["GOZUID"]].value))
    Organization.set("ContractDate",str(current_datetime(sheet1[config["start_indexes"]["ContractDate"]].value))) 
                                                 
    Forms=ET.SubElement(root,"Forms")
    Form8=ET.SubElement(Forms,"Form",{'type':'8',"ReportDate":str(now_local),'Year':str(sheet1[config['start_indexes']['Year']].value),
                                      'Quarter':str(sheet1[config["start_indexes"]["Ouarter"]].value)})
    ContractSpending=ET.SubElement(Form8,"ContractSpending")
    Contractors = ET.SubElement(Form8, 'Contractors')
    ContractFinance=ET.SubElement(Form8,"ContractFinance")
    for row in range(find_first_row_with_prefix(sheet1,1,"1."), sheet1.max_row + 1):
      if str(sheet1.cell(row=row, column=1).value).startswith('1.1'):  
        ContractSpending.set("Salary",str(int(sheet1.cell(row=row, column=3).value*100)))
      if str(sheet1.cell(row=row, column=1).value).startswith('1.2'):  
        ContractSpending.set("Taxes",str(int(sheet1.cell(row=row, column=3).value*100)))
      if str(sheet1.cell(row=row, column=1).value).startswith('1.3'):   
        ContractSpending.set("Rates",str(int(sheet1.cell(row=row, column=3).value*100)))
      if str(sheet1.cell(row=row, column=1).value).startswith('1.4'):    
        ContractSpending.set("Other",str(int(sheet1.cell(row=row, column=3).value*100)))
      if str(sheet1.cell(row=row, column=1).value).startswith('1.5'):  
        ContractSpending.set("Reserve",str(int(sheet1.cell(row=row, column=3).value*100)))
      if str(sheet1.cell(row=row, column=1).value).startswith('1.6'):  
        ContractSpending.set("Income",str(int(sheet1.cell(row=row, column=3).value*100)))
      if str(sheet1.cell(row=row, column=1).value).startswith('2.'):
       Contractor = ET.SubElement(Contractors,"Contractor",{'Total':str(int(sheet1.cell(row=row, column=3).value)*100),
                                                           'Name':str(sheet1.cell(row=row, column=4).value),
                                                           'INN':str(sheet1.cell(row=row, column=5).value),
                                                           'ContractNumber':str(sheet1.cell(row=row, column=6).value),
                                                           'ContractDate': str(current_datetime(sheet1.cell(row=row, column=7).value)),
                                                           'AccountNumber':str(sheet1.cell(row=row, column=8).value),
                                                           'Cost': str(int(sheet1.cell(row=row, column=9).value)*100),
                                                           'PaymentPlanned':str(int(sheet1.cell(row=row, column=10).value)*100),
                                                           'FinishDate': str(current_datetime(sheet1.cell(row=row, column=12).value))})
        
    #nextCell=find_last_row_with_prefix(sheet1,1,'2')
    if str(sheet1.cell(row=row, column=1).value).startswith('3'):
     PlannedPay=ET.SubElement(Form8,"PlannedPay",{"Total":str(int(sheet1.cell(row=row, column =9).value)*100), #цена договора?
                                                 "PaymentPlanned":str(int(sheet1(row=row, column =10).value)*100),
                                                 "PaymentCurrent":str(int(sheet1(row=row, column =3).value)*100)})
    if str(sheet1.cell(row=row, column=1).value).startswith('4'):
     ContractFinance.set("TotalRequirement",str(int(sheet1.cell(row=row, column=3).value*100)))
    if str(sheet1.cell(row=row, column=1).value).startswith('5'): 
     ContractFinance.set("CashBalance",str(int(sheet1.cell(row=row, column=3).value)*100))
    if str(sheet1.cell(row=row, column=1).value).startswith('6'):
     ContractFinance.set("PlannedIncome",str(int(sheet1.cell(row=row, column=3).value)*100))
    if str(sheet1.cell(row=row, column=1).value).startswith('7'):
     ContractFinance.set("DepositeIncome",str(int(sheet1.cell(row=row, column=3).value)*100))
    
    FormSupplement=ET.SubElement(Forms,"Form",{'type':'Supplement',"ReportDate":str(now_local)})
    Parts=ET.SubElement(FormSupplement,"Parts")
   # хз будем считать что с 3 строки и год начинается с 202 и данные начинаются с 1 ячейки
    for row in range(3, sheet1.max_row + 1):
      if str(sheet2.cell(row=row, column=1).value).startswith('202'):
        Part=ET.SubElement(Parts,"Part",{"Year":str(sheet2.cell(row=row, column=1).value),
                                  "Quarter":str(sheet2.cell(row=row, column=2).value),
                                  "Requirement":str(int(sheet2.cell(row=row, column=3).value)*100),
                                  "Deviation":str(int(sheet2.cell(row=row, column=4).value)*100)})
       
        reasons = str(sheet2.cell(row=row, column=5).value).split(',')
        if any(reason.strip() != 'None' and reason.strip() != '' for reason in reasons):
           Reasons = ET.SubElement(Part,"Reasons")
        for reason in reasons:
         if  reason.strip() != 'None' and reason.strip() != '':
            ET.SubElement(Reasons,"Reason").text=reason.strip()
    xmlstr = xml_to_pretty_string(root)
    
    return xmlstr,config

def update_config(filename, config):
    with open(filename, "w", encoding='utf-8') as configfile:
        config_json = json.dumps(config, ensure_ascii=False)
        configfile.write(config_json)

def save_xml(xmlstr, filename):
    with open(filename, "wb") as f:
        f.write(xmlstr)

def main():
    root = tk.Tk()
    root.withdraw()
    
    from tkinter import filedialog
    xlsx_file_path = filedialog.askopenfilename(title="Выберите XLSX файл", filetypes=[("XLSX Files", "*.xlsx")])
    config = load_config("config.json")
    workbook = load_workbook_data(xlsx_file_path)
    sheet1 = workbook.worksheets[0]
    sheet2 = workbook.worksheets[1]
    xmlstr,updated_config = create_xml(sheet1,sheet2, config, current_datetime())
    save_xml(xmlstr, os.path.dirname(xlsx_file_path) + "\message.xml")
    update_config("config.json", updated_config) 
    add_to_zip(os.path.dirname(xlsx_file_path)+'\message.xml',config["INN"],config["PreviousUID"])
    messagebox.showinfo("SAIBIS","XML отчет сохранен рядом с " + xlsx_file_path)
    
main()
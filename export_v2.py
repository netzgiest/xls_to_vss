from openpyxl import load_workbook
from xml.dom.minidom import parseString
import xml.etree.ElementTree as ET
from datetime import datetime
import json,os,pytz, uuid,zipfile
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as messagebox

def load_config(filename):
    with open(filename, "r", encoding='utf-8') as configfile:
        config_json = configfile.read()
    config = json.loads(config_json)
    return config

def current_datetime(value=None):
    tz = pytz.timezone("Europe/Moscow")
    if value is None:
     now_utc = datetime.now(pytz.UTC)
     now_local = now_utc.astimezone(tz)
    else:
     now_local= tz.localize(datetime.strptime(str(value), "%d.%m.%Y")).replace(hour=0, minute=0, second=0)  
    return now_local.isoformat()

def load_workbook_data(filename):
    wb = load_workbook(filename=filename, read_only=True)
    return wb

def prettify(elem):
    from xml.etree.ElementTree import tostring
    rough_string = tostring(elem, 'utf-8')
    reparsed = parseString(rough_string)
    return reparsed.toprettyxml(indent="   ", encoding="UTF-8")

def add_to_zip(filepath,INN,UID):
    zip_name = f"{INN}_{datetime.now().strftime('%Y%m%d')}_{UID}.zip"

    with zipfile.ZipFile(zip_name,'w') as zipf:
       zipf.write(filepath)
       return filepath 

def create_xml(sheet1,sheet2, config, now_local):
    thisUID = str(uuid.uuid4())
    root =ET.Element('Message',{'CreateDate':str(now_local),'UID':thisUID,'PreviousUID':str(config["PreviousUID"])})
    config["PreviousUID"] = thisUID
    Organization =ET.SubElement(root, "Organization",{'INN':str(config["INN"]),"KPP":config["KPP"],'Name':config["NAME"],
                                                     "GOZUID":str(sheet1[config["start_indexes"]["GOZUID"]].value),
                                                     "ContractDate":str(current_datetime(sheet1[config["start_indexes"]["ContractDate"]].value)) 
                                                     })
    Forms = ET.SubElement(root,"Forms")
    Form8=ET.SubElement(Forms,"Form",{'type':'8',"ReportDate":str(now_local),'Year':str(sheet1[config['start_indexes']['Year']].value),
                                      'Quarter':str(sheet1[config["start_indexes"]["Ouarter"]].value)})
    ContractSpending=ET.SubElement(Form8,"ContractSpending",{"Salary":str(sheet1[config["start_indexes"]["Salary"]].value*100),
                                                             "Taxes":str(sheet1[config["start_indexes"]["Taxes"]].value*100),
                                                             "Rates":str(sheet1[config["start_indexes"]["Rates"]].value*100),
                                                             "Other":str(sheet1[config["start_indexes"]["Other"]].value*100),
                                                             "Reserve":str(sheet1[config["start_indexes"]["Reserve"]].value*100),
                                                             "Income":str(int(sheet1[config["start_indexes"]["Income"]].value)*100)
                                                             })
    Contractors = ET.SubElement(Form8, 'Contractors') 
    nextCell = int(config["start_indexes"]["startcontractor"])
    for row in range(nextCell, sheet1.max_row + 1):
     if str(sheet1.cell(row=row, column=1).value).startswith('2'):
      Contractor = ET.SubElement(Contractors,"Contractor",{'Total':str(int(sheet1.cell(row=row, column=3).value)*100),
                                                           'Name':str(sheet1.cell(row=row, column=4).value),
                                                           'INN':str(sheet1.cell(row=row, column=5).value),
                                                           'ContractNumber':str(sheet1.cell(row=row, column=6).value),
                                                           'ContractDate': str(current_datetime(sheet1.cell(row=row, column=7).value)),
                                                           'AccountNumber':str(sheet1.cell(row=row, column=8).value),
                                                           'Cost': str(int(sheet1.cell(row=row, column=9).value)*100),
                                                           'PaymentPlanned':str(int(sheet1.cell(row=row, column=10).value)*100),
                                                           'FinishDate': str(current_datetime(sheet1.cell(row=row, column=12).value))})
        
      nextCell=row+1
    PlannedPay=ET.SubElement(Form8,"PlannedPay",{"PaymentPlanned":str(int(sheet1['J'+str(nextCell)].value)*100),
                                                 "Total":str(int(sheet1['I'+str(nextCell)].value)*100),
                                                 "PlannedPay":str(int(sheet1['C'+str(nextCell)].value)*100)})
    ContractFinance=ET.SubElement(Form8,"ContractFinance",{"TotalRequirement":str(int(sheet1['C'+str(nextCell+1)].value)*100),
                                                           "CashBalance":str(int(sheet1['C'+str(nextCell+2)].value)*100),
                                                           "PlannedIncome":str(int(sheet1['C'+str(nextCell+3)].value)*100),
                                                           "DepositeIncome":str(int(sheet1['C'+str(nextCell+4)].value)*100)})
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
           Reasons = ET.SubElement(Parts,"Reasons")
        for reason in reasons:
         if  reason.strip() != 'None' and reason.strip() != '':
            ET.SubElement(Reasons,"Reason").text=reason.strip()
    xmlstr = prettify(root)
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
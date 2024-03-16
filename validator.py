from tkinter import filedialog
import tkinter.messagebox as messagebox

from lxml import etree
xlsx_file_path = filedialog.askopenfilename(title="Выберите Xml файл", filetypes=[("xml Files", "*.xml")])
try:
      with open('message.xsd', 'rb') as f:
       schema_root = etree.XML(f.read())
       schema = etree.XMLSchema(schema_root)
      try:     
       with open(xlsx_file_path, 'rb') as f:
        xml_root = etree.XML(f.read())
       try:
        schema.assertValid(xml_root)
        messagebox.showinfo("saibis","XML соответствует XSD схеме.")
       except etree.DocumentInvalid as err:
        messagebox.showinfo( "Ошипка","Ошибка валидации:"+ str(err))
      except IOError as err:
       print(err) 
except IOError as err:
         print(err)
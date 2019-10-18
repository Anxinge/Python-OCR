# -*- coding: utf-8 -*-
"""
Created on Mon Mar 25 15:12:49 2019

@author: ashwin
"""

import PySimpleGUI as sg
import re
import datetime
import pyap
try:
    import Image
except ImportError:
    from PIL import Image
import pytesseract 
import difflib
import json
#import dateutil.parser as dparser


tesseractpath="C:/Users/ashwin/Anaconda3/Lib/site-packages/"

resultfileName = '/Result_Analysing'
sg.SetOptions(
                 button_color = sg.COLOR_SYSTEM_DEFAULT
               , text_color = sg.COLOR_SYSTEM_DEFAULT
             )
sg.ChangeLookAndFeel('GreenTan')      

layout=[[sg.Text('1.InvoiceFilename', size=(15, 1),auto_size_text=False, justification='left')
        ,sg.Input(key='_IN_',size=(43,1)),sg.FileBrowse('Select_Invoice', button_color=('white', 'green'),size=(15,1))],
        [sg.Text('2.Add to FileList', size=(15, 1), auto_size_text=False, justification='left'),
        sg.Listbox(values=['invoice1.jpg','invoice2.jpg','...'],select_mode='multiple',size=(41, 6), key='_FileList_')
         ,sg.Button('Add_File_List', button_color=('white', 'green'),size=(15, 1))],
        [sg.Text('3.Select Result Folder', size=(15, 1),  auto_size_text=False, justification='left'), 
         sg.InputText('Result_Folder',key='_ResultFolder_',size=(43,1))
         ,sg.FolderBrowse('Result Folder',button_color=('white', 'green'),size=(15,1))],
        [sg.Text('4.Extract Text', size=(15, 1),  auto_size_text=False, justification='left')
        ,sg.OK('Detect Text', button_color=('white', 'green'),size = (15,1)),sg.Text('5.Analysing Text     ', size=(15, 1))
        ,sg.OK('Analysing', button_color=('white', 'green'),size = (15,1))],
        [sg.Multiline(default_text='Extract Text from Imagefiles',size=(37,15),key='_DetectText_'),
         sg.Multiline(default_text='Analysing Text Result',size=(37,15),key='_FinalResult_')],
        [sg.Cancel()] ]
window = sg.Window('Invoice Image Analysis').Layout(layout) 

sample_text = ['Invoice','INVOICE','Importer Name','JAPAN','details','Account','Quantity','Goods',
               'Credit','TOTAL','Total','COMPANY','Bill','From','Phone Number']
country_classes =['China','Japan','Italy','USA','French','Hong Kong','United States','Poland',
                'United KingDom','Saudi Arabia','Singapore','South Korea','England','Swiss']

def strToTxt(resultfileName,out_Text):
    with open(resultfilepath + resultfileName + '.txt','a',encoding='utf-8') as f:
        f.write(out_Text)
        f.write('\n')
        
List_File = []
while True:                 # Event Loop  
  event, values = window.Read()
  #print(event,values)
  #imagefilepath = values[0]
  #print(imagefilepath)
  if event is None or event == 'Exit':  
      break 
  if event == 'Add_File_List':
      List_File.append(str(values['_IN_']))
      window.FindElement('_FileList_').Update(List_File)
  
      
      
  if event == 'Detect Text':
      num_Files = len(List_File)
      out_text_total = ""
      for i in range(num_Files):
          #print(values)
          imagefilepath = List_File[i]
          print(imagefilepath)
          resultfilepath = values['_ResultFolder_']
          print(resultfilepath)
          #pytesseract.pytesseract.tesseract_cmd = tesseractpath
          ### open invoice image as a input
          image = Image.open(imagefilepath)
          # Simple image to string
          out_text=pytesseract.image_to_string(image, lang = 'eng')
          out_text="".join([s for s in out_text.strip().splitlines(True) if s.strip()])
          ##  print the extract txt
          print(out_text)
          out_text_total += out_text 
          window.FindElement('_DetectText_').Update(out_text_total)
          
      
  if event == 'Analysing':
      list_text = out_text_total.splitlines()
      len_line = len(list_text)
      for k in range(len_line):
          list_text[k]=re.sub('[=._..__"^‘`*]',' ',list_text[k])
      pipeline = []
      
      country_list=[]
      adress_list = []
      for country in country_classes:
          for line_num in range(len_line):
              words=list_text[line_num].split(' ')
              if country in words:
                  country_list.append(country)
      pipeline.append('Country : ' + str(country_list))
      adress = pyap.parse(out_text_total,country='US')
      
      if len(adress)!=0:
          pipeline.append('ImporterAdress : ' + str(adress))
      date_list=[]
      for line_num in range(len_line):

          try:
              match = re.search(r'\d{4}-\d{2}-\d{2}',list_text[line_num] )
          except:
              pass
          if match:
              date = datetime.datetime.strptime(match.group(), '%Y-%m-%d').date()
              date_list.append(str(date))
          try:
              match_4 = re.search(r'\d{4}-\d{1}-\d{1}',list_text[line_num] )
          except:
              pass
          if match_4:
              date = datetime.datetime.strptime(match_4.group(), '%Y-%m-%d').date()
              date_list.append(str(date))
          try:
              match_1 = re.search(r'\d{4}/\d{2}/\d{2}',list_text[line_num] )
          except:
              pass
          if match_1:
              date = datetime.datetime.strptime(match_1.group(), '%Y/%m/%d').date()
              date_list.append(str(date))
          try:
              match_2 = re.search(r'\d{1}/\d{1}/\d{4}',list_text[line_num] )
          except:
              pass
          if match_2:
              date = datetime.datetime.strptime(match_2.group(), '%d/%m/%Y').date()
              date_list.append(str(date))
          try:
              match_3 = re.search(r'\d{2}/\d{2}/\d{4}',list_text[line_num] )
          except:
              pass
          if match_3:
              date = datetime.datetime.strptime(match_1.group(), '%d/%m/%Y').date()
              date_list.append(str(date))
      pipeline.append('Datetime : ' + str(date_list))                
      for txt in sample_text:
          txt_value_list=[]
          for line_num in range(len_line):

              words=list_text[line_num].split(' ')
              result=difflib.get_close_matches(txt,words)
              if len(result)!=0 and len(words) >= 2 and list_text[line_num]!="COMMERCIAL INVOICE":
                  key = txt
                  txt_value_list.append(list_text[line_num])
          if len(txt_value_list) != 0:
              pipeline.append(txt + ' : ' + str(txt_value_list))
      window.FindElement('_FinalResult_').Update(pipeline)
# =============================================================================
#                   if txt == 'INVOICE' or txt == 'Invoice':
#                       print(txt)
#                       key = 'Invoice Number'
#                       result = re.findall(r"\d+",list_text[line_num])  
#                       if result:
#                           value=result.group()
#                           pipeline.append((key + ':' + value.strip()))
#                   if txt == 'Phone Number':
#                       phoneNumRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d\d')
#                       mo = phoneNumRegex.search(list_text[line_num])
#                       key = txt
#                       value = mo.group
#                       pipeline.append((key + ':' + value.strip()))
#                   if txt == 'Importer Name':
#                       nameRegex = re.compile(r'Co Ltd: (.*)')
#                       mo = nameRegex.search(list_text[line_num])
#                       key = txt
#                       value = mo.group
#                       pipeline.append((key + ':' + value.strip()))
#       window.FindElement('_FinalResult_').Update(pipeline)
# =============================================================================
      
      
      filePathNameWExt =  resultfilepath  + resultfileName + '.json'
      with open(filePathNameWExt, 'a') as f:
          json.dump(pipeline, f)
      
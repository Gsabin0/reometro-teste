import PyPDF2 
import re
from datetime import datetime
import pyodbc 
import os
import win32com.client as win32
import pyautogui as pg
import time


pastapara = "Q:\\Docs Laboratorio\\CURVAS REOMETRO\\Arquivos_Concluidos\\"
dt = datetime.today().strftime('%Y-%m-%d') 
print(dt)
      
try:
    server = '192.168.0.30' 
    database = 'ELIPSE_E3' 
    username = 'elipse' 
    password = 'E#lipse#365#ic' 
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';ENCRYPT=no;UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    print("deu certo")
except:
    print(("Erro na conexão com o banco de dados"))

try:    
    os.chdir("Q:\\Docs Laboratorio\\CURVAS REOMETRO")#("C:\\Users\\tom\\Documents\\testepy")
    todos_arquivos = os.listdir()
except:
    print("erro na pasta do reometro")

try:
    for arquivos in todos_arquivos:
        if ".pdf" in arquivos:
            todos_pdf = arquivos
            #ler pdf e extraix lote
            pdfdoc = open(f'{todos_pdf}', 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfdoc)
            pageObj = pdfReader.getPage(0)
            pdf_extra = pageObj.extractText()
            
            mh = re.compile(r' ((\.|,|[0-9])*) (?:\.|,|[0-9])*  ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (.|\s)(?:\.|,|[0-9])*  (?:\.|,|[0-9])*  (?:\.|,|[0-9])*  (?:\.|,|[0-9])*  (?:\.|,|[0-9])* .ote (([0-9]+)) (([0-9]+))(\w+)')
            check_finditer = mh.finditer(pdf_extra)
            complet_pdf = check_finditer
            _mh = mh.finditer(pdf_extra)
                
            for _mh in complet_pdf:
                    print(_mh)
                    print(_mh.group(1))
                    print(_mh.group(3))
                    print(_mh.group(5))
                    print(_mh.group(7))
                    print(_mh.group(10))
                    print(_mh.group(13))
                    print(_mh.group(14))
                    ML = float(_mh.group(1).replace(',', '.'))
                    MH = float(_mh.group(3).replace(',', '.'))
                    TS2 = float(_mh.group(5).replace(',', '.'))
                    T90 = float(_mh.group(7).replace(',', '.'))
                    lot = _mh.group(10)
                    amos = _mh.group(13)
                    suf = _mh.group(14)

                    #composto
                    composto = re.compile(r'\d{10}')
                    check_finditer = composto.finditer(pdf_extra)
                    complet_pdf = check_finditer
                    _composto = composto.finditer(pdf_extra)
                    for _composto in complet_pdf:
                        print(_composto)
                        print(_composto.group(0))
                        composto = _composto.group(0)
                
                #ml mh ts2 t90
                    
                
                    cursor.execute(f"""INSERT INTO [dbo].[REOMETRO_COPIA]
                                ([COMPOSTO]
                                ,[Lote]
                                ,[Amos]
                                ,[Suf]
                                ,[ML]
                                ,[MH]
                                ,[ts2]
                                ,[T90]
                                ,[Data]
                                ,[ID_Setup_GPM])
                                VALUES

                            ('{composto}','{lot}','{amos}','{suf}',{ML},{MH},{TS2},{T90},'{dt}',NULL)""")

                    cnxn.commit()
                    pdfdoc.close()  
    
                  

    

except:
    print("erro ao gravar no banco")  

#os.chdir(r"Q:\Docs Laboratorio\CURVAS REOMETRO")
#todos_arquivos = os.listdir()
#for arquivos in todos_arquivos:
#        if ".pdf" in arquivos:
#            todos_pdf = arquivos
#            os.rename(todos_pdf, pastapara+todos_pdf) 
#


                 
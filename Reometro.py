import PyPDF2 
import re
from datetime import datetime
import pyodbc 
import os


dt = datetime.today().strftime('%Y-%m-%d') 
print(dt)
      
server = '192.168.0.30' 
database = 'ELIPSE_E3' 
username = 'elipse' 
password = 'E#lipse#365#ic' 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';ENCRYPT=no;UID='+username+';PWD='+ password)

print("deu certo")
cursor = cnxn.cursor()

os.chdir(r"C:\Users\tom\Documents\testepy")
todos_arquivos = os.listdir()
contador = 1

for arquivos in todos_arquivos:
    if ".pdf" in arquivos:
        todos_pdf = arquivos
        #ler pdf e extraix lote
        pdfdoc = open(f'{todos_pdf}', 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfdoc)
        pageObj = pdfReader.getPage(0)
        pdf_extra = pageObj.extractText()

        lote = re.compile(r'.ote (([0-9]+)) (([0-9]+))(\w+)')
        check_finditer = lote.finditer(pdf_extra)
        complet_pdf = check_finditer
        _lote = lote.finditer(pdf_extra)

        for _lote in complet_pdf:
                
            print(_lote)
            print(_lote.group(2))
            print(_lote.group(3))
            print(_lote.group(5))
            lot = _lote.group(2)
            amos = _lote.group(3)
            suf = _lote.group(5)

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
                
                
                        
                mh = re.compile(r'.rados(.|\s) ((\.|,|[0-9])*) (?:\.|,|[0-9])*  ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (?:\.|,|[0-9])* (.|\s)(?:\.|,|[0-9])*  (?:\.|,|[0-9])*  (?:\.|,|[0-9])*  (?:\.|,|[0-9])*  (?:\.|,|[0-9])* .ote [0-9]+ [0-9a-zA-Z]+ [0-9]+/[a-zA-Z]+/[0-9]+  (?:\.|:|[0-9])* [A-Za-z]+ [a-zA-Z]+ [a-zA-Z]+ SI(.|\s) ((\.|,|[0-9])*) (?:\.|,|[0-9])*  ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* ((\.|,|[0-9])*) (?:\.|,|[0-9])* (?:\.|,|[0-9])* ((\.|,|[0-9])*) (?:\.|,|[0-9])* ((\.|,|[0-9])*)')
                check_finditer = mh.finditer(pdf_extra)
                complet_pdf = check_finditer
                _mh = mh.finditer(pdf_extra)
                for _mh in complet_pdf:
                    print(_mh)
                    print(_mh.group(2))
                    print(_mh.group(4))
                    print(_mh.group(6))
                    print(_mh.group(8))

                    ML = float(_mh.group(2).replace(',', '.'))
                    MH = float(_mh.group(4).replace(',', '.'))
                    TS2 = float(_mh.group(6).replace(',', '.'))
                    T90 = float(_mh.group(8).replace(',', '.'))

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
                    

                
                       
               
                
       
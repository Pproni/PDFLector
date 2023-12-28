from spire.pdf.common import *
from spire.pdf import *
import pdfplumber
import pandas as pd
from openpyxl import load_workbook, Workbook
import re
import os
import numpy as np
from datetime import datetime, timedelta

def folder_creator(folder_directory,folder_name):
    path = os.path.join(folder_directory, folder_name)
    if not os.path.exists(path):
        os.mkdir(path)

def check_pdfs(path):
    files = os.listdir(path)
    spliting = [element.split('.', 1) for element in files]
    files_names = [spliting[i][0] for i in range(len(spliting))]
    return files, files_names

def pdf2svg(pdf,pdf_name,x=1800,y=1600):
    
    path_pdf = os.path.join(str(os.getcwd()),'PDFs')
    path_svg = os.path.join(str(os.getcwd()),'SVGs')
    path_data = os.path.join(str(os.getcwd()),'data')
    
    # Create an object of the PdfDocument class
    doc = PdfDocument()

    # Load a PDF file
    doc.LoadFromFile(os.path.join(path_pdf,str(pdf)))

    # Specify the width and height of output SVG files
    doc.ConvertOptions.SetPdfToSvgOptions(float(x), float(y))

    # Save each page of the file to a separate SVG file
    doc.SaveToFile(os.path.join(path_svg,str(pdf_name))+".svg", 0, 0, FileFormat.SVG)

    # Create a PdfDocument object
    doc1 = PdfDocument()
    # Load an SVG file
    doc1.LoadFromSvg(os.path.join(path_svg,str(pdf_name))+".svg")

    # Save the SVG file to PDF format
    doc1.SaveToFile(os.path.join(path_data,str(pdf_name))+".pdf", FileFormat.PDF)

    # Close the PdfDocument object
    doc.Close()
    doc1.Close()
    
def pdf2excel(pdf_file,pdf_name):
    path_data = os.path.join(str(os.getcwd()),'data')
    path_excel = os.path.join(str(os.getcwd()),'Excel')
    
    pdf = pdfplumber.open(os.path.join(path_data,str(pdf_file)))
    p0 = pdf.pages[0]
    table = p0.extract_table()

    # Crear DataFrame
    df = pd.DataFrame(table[1:])

    # Invertir las filas
    df = df[::-1].reset_index(drop=True)

    # Guardar en Excel
    df.to_excel(os.path.join(path_excel,str(pdf_name))+".xlsx", index=False)

def data_extractor(excel_file):
    path_excel = os.path.join(str(os.getcwd()),'Excel')
    workbook = load_workbook(filename=os.path.join(path_excel,str(excel_file))+".xlsx")
    sheet = workbook.active
    #Extracción de nombres, horas totales, hora de inicio, final, y descanso
    names = []
    total_hours = []
    time_list = []
    
    for i in range(14):
        j = 2*i+17
        cells = sheet[str("H"+str(j)): str("J"+str(j))]
        if str(sheet[str("B"+str(j))].value).replace(" ","") == str(707):
            #Nombres
            names.append(sheet[str("C"+str(j))].value)
            #Horas totales
            total_hours.append((sheet[str("L"+str(j-1))].value).replace(" ",""))
            #Horas iniciales, descanso y finales
            if str(sheet[str("I"+str(j))].value).replace(" ","") == str(0):
                clock_list = [time_list.append([str(c1.value),str(0) ,str(c2.value),str(c3.value)]) for (c1,c2,c3) in cells if str(c1.value) != 'None']
                time_list = [[h1.replace(" ",""),str(h2),h3.replace(" ",""),h4.replace(" ","")] for x in time_list for (h1,h2,h3,h4) in [x]]
                time_list = list(filter(None,time_list))                
            elif str(sheet[str("I"+str(j))].value).replace(" ","") != str(0):
                clock_list = [time_list.append([str(c1.value),str(sheet[str("I"+str(j-1))].value) ,str(c2.value),str(c3.value)]) for (c1,c2,c3) in cells if str(c1.value) != 'None']
                time_list = [[h1.replace(" ",""),str(h2).replace(" ",""),h3.replace(" ",""),h4.replace(" ","")] for x in time_list for (h1,h2,h3,h4) in [x]]
                time_list = list(filter(None,time_list))
            else:
                print('Hay un dato atípico en la fila: '+j)
                pass
    #Extracción de Trabajo, código de trabajo, fecha de inicio y final
    job_ID_text = re.split('(\d+)',str(sheet[str("F4")].value).replace(" ",""))
    job_name = (str(sheet[str("F7")].value).replace(" ","")).replace("\n","")
    start_day = (str(sheet[str("B5")].value).replace(" ","")).replace("STARTDATE","").replace("\n","")
    end_day = (str(sheet[str("B8")].value).replace(" ","")).replace("STOPDATE","").replace("\n","")
    total_hours = [float(x) for x in total_hours]
    names = [x.replace(" ", "") for x in names]
    for i in job_ID_text:
        if i.isdigit():
            job_ID = i        
    return names, job_name.replace("JOBNAME",""), job_ID, total_hours, time_list, start_day, end_day

def daylist_generator(initial_date, final_date):
    # Convierte las cadenas de fecha en objetos datetime
    initial_date = datetime.strptime(initial_date, "%Y-%m-%d")
    final_date = datetime.strptime(final_date, "%Y-%m-%d")

    # Calcula la diferencia entre las fechas
    diff = final_date - initial_date

    # Crea una lista de días
    day_list = [initial_date + timedelta(days=d) for d in range(diff.days + 1)]
    day_list = [(str(dia.strftime('%A'))+" "+ str(dia.day)) for dia in day_list]

    return day_list

def excel_creator(name_list, date_list): #FALTA AGREGAR LAS HORAS REGULARES Y OVERTIME DE LA PRIMERA TABLA, POR DÍA.
    
    def get_excel_column_name(column_number):
        """
        Convierte el número de columna al estilo de nombre de columna de excel.
        Ej: 1->A, 2->B, 26->Z, 27->AA, etc.
        """
        result = ""
        while column_number > 0:
            remainder = (column_number - 1) % 26
            result = chr(65 + remainder) + result
            column_number = (column_number - 1) // 26
        return str(result)
    
    # Crear un libro de trabajo
    workbook = Workbook()

    # Crear una hoja con título específico
    #workbook.create_sheet(title='General')
    sheet = workbook['Sheet']
    sheet.title = 'General'
    
    #Agregar las fechas de inicio y final
    sheet['A1'] = 'Start day of week: '+ str(date_list[0])+"\n"+' End day of week: '+ str(date_list[-1])
    #sheet['B1'] = 'End day of week: '+ str(date_list[-1])
    
    #Agregar nombres
    espacio_pagos = 1
    
    sheet['A2'] = 'NAMES'
    sheet[get_excel_column_name(36 + espacio_pagos) + '2'] = 'NAMES'
    for i in range(len(name_list)):
        sheet['A'+str(i+3)] = name_list[i]
        sheet[get_excel_column_name(36 + espacio_pagos) + str(i+3)] = name_list[i]
    
    #Agregar los días de trabajo
    for i in range(len(date_list)):
        sheet[str(chr(66+i))+'2'] = date_list[i]
        
    #Agregar celdas con horas diarias
    cell_text_perday = ['TOTAL HOURS DAY - DAILY', 'TOTAL REGULAR HOURS - DAILY', 'TOTAL OVERTIME HOURS - DAILY']
    for i in range(len(cell_text_perday)):
        sheet['A'+str(len(name_list)+3+i)] = cell_text_perday[i]
        for j in range(len(date_list)):
            #Para este caso se tiene que editar pensando en que las celdas tienen valores dados otras tablas según la plantilla que se usa generalmente, es decir,
            #los valores de la casilla Total Regular Hours - Daily debe variar si hay overtime en alguna de las casillas
            if i==0:
                sheet[str(chr(66+j))+str(len(name_list)+3+i)] = "=SUM("+str(chr(66+j))+'3:'+str(chr(66+j))+str(len(name_list)+2)+')'
            elif i==1:
                sheet[str(chr(66+j))+str(len(name_list)+3+i)] = f"={get_excel_column_name(21+espacio_pagos+j)}{len(name_list)+3}"
            elif i==2:
                sheet[str(chr(66+j))+str(len(name_list)+3+i)] = f"={get_excel_column_name(29+espacio_pagos+j)}{len(name_list)+3}"
    
    #Formato y código de las horas semanales            
    cell_text_week = ['TOTAL HOURS - WEEKLY', 'TOTAL REGULAR HOURS - WEEKLY', 'TOTAL OVERTIME HOURS - WEEKLY']
    for i in range(len(cell_text_week)):
        sheet[str(chr(66+len(date_list)+i))+'2'] = cell_text_week[i]
        for j in range(len(name_list)):
            if i==0:
                sheet[str(chr(66+len(date_list)+i))+str(3+j)] = "=SUM("+str(chr(66))+str(3+j)+':'+str(chr(66+len(date_list)-1))+str(3+j)+')'
            elif i==1:
                sheet[str(chr(66+len(date_list)+i))+str(3+j)] = "=IF(I"+str(3+j)+'<=40,I'+str(3+j)+',40)'
            elif i==2:
                sheet[str(chr(66+len(date_list)+i))+str(3+j)] = "=I"+str(3+j)+"-J"+str(3+j)
    
    #Realiza la suma de cada una de las columnas
    sheet['I13'] = '=SUM(I3:I'+str(len(name_list)+2)+')'
    sheet['J14'] = '=SUM(J3:J'+str(len(name_list)+2)+')'
    sheet['K15'] = '=SUM(K3:K'+str(len(name_list)+2)+')'
    
    #Llena las celdas vacías
    for fila in range(int(len(name_list)+3),int(len(name_list)+6)):
        for columna in range(9,12):
            if sheet.cell(row=fila, column=columna).value == None:
                sheet.cell(row=fila, column=columna, value='-')
            else:
                pass
            
    #Tablas extra
    ##Total Hours table
    ###Nombres de las tablas
    sheet[str(chr(77+espacio_pagos))+'1'] = 'TOTAL HOURS (ACCUMULATED)'
    sheet[get_excel_column_name(21+espacio_pagos)+'1'] = 'REGULAR HOURS'
    sheet[get_excel_column_name(29+espacio_pagos)+'1'] = 'OVERTIME HOURS (PER DAY)'
    
    ###Agregar los días de trabajo y horas totales acumuladas
    for i in range(len(date_list)):
        sheet[str(chr(77+i+espacio_pagos))+'2'] = date_list[i]
        for j in range(len(name_list)):
            if i==0:
                sheet[str(chr(77+i+espacio_pagos))+str(j+3)] = f"=B{j + 3}"
            else:
                previous_cell = str(chr(77+i+espacio_pagos-1))+str(j+3)
                general_table_cell = str(chr(66+i))+str(j+3)
                sheet[str(chr(77+i+espacio_pagos))+str(j+3)] = f"={general_table_cell}+{previous_cell}"
    
    ##Regular Hours
    ###Agregar los días de trabajo y horas totales acumuladas
    for i in range(len(date_list)):
        sheet[get_excel_column_name(21 + i + espacio_pagos) + '2'] = date_list[i]
        for j in range(len(name_list)):
            if i == 0:
                sheet[get_excel_column_name(21 + i + espacio_pagos) + str(j + 3)] = f"=N{j + 3}"
            else:
                sheet[get_excel_column_name(21 + i + espacio_pagos) + str(j + 3)] = f"=IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}<=0, 0, IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}<=40,{get_excel_column_name(13 + espacio_pagos + i)}{j+3}-{get_excel_column_name(13 + espacio_pagos + i - 1)}{j+3},IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}-{get_excel_column_name(13 + espacio_pagos + i-1)}{j+3}<=0, 0, ABS({get_excel_column_name(13 + espacio_pagos + i)}{j+3}-{get_excel_column_name(13 + espacio_pagos + i-1)}{j+3}-{get_excel_column_name(29 + espacio_pagos + i)}{j+3}))))"

    ##Overtime Hours
    for i in range(len(date_list)):
        sheet[get_excel_column_name(29 + i + espacio_pagos) + '2'] = date_list[i]
        for j in range(len(name_list)):
            if i == 0:
                sheet[get_excel_column_name(29 + i + espacio_pagos) + str(j + 3)] = "=0"
            else:
                sheet[get_excel_column_name(29 + i + espacio_pagos) + str(j + 3)] =f"=IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}<=0, 0, IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}<=40,0, IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}-{get_excel_column_name(13 + espacio_pagos + i-1)}{j+3}<=0,0,IF({get_excel_column_name(13 + espacio_pagos + i)}{j+3}>40, {get_excel_column_name(13 + espacio_pagos + i)}{j+3}-40-SUM({get_excel_column_name(29 + espacio_pagos)}{j+3}:{get_excel_column_name(29 + espacio_pagos + i - 1)}{j+3}),0))))"
    
    for i in range(len(date_list)):
        sheet[chr(77+i+espacio_pagos)+str(len(name_list)+3)] = f"=SUM({chr(77+i+espacio_pagos)}3:{chr(77+i+espacio_pagos)}{len(name_list)+2})"
        sheet[get_excel_column_name(21 + i + espacio_pagos) + str(len(name_list) + 3)] = f"=SUM({get_excel_column_name(21 + i + espacio_pagos)}3:{get_excel_column_name(21 + i + espacio_pagos)}{len(name_list) + 2})"
        sheet[get_excel_column_name(29 + i + espacio_pagos) + str(len(name_list) + 3)] = f"=SUM({get_excel_column_name(29 + i + espacio_pagos)}3:{get_excel_column_name(29 + i + espacio_pagos)}{len(name_list) + 2})"
       
    # Guardar el libro de trabajo en un archivo
    workbook.save('test_savedata.xlsx')
    workbook.close()

def new_sheets(name_list, job_name, job_ID, date_list, job_day, job_names, hours):
    workbook = load_workbook('test_savedata.xlsx')
    if not str(job_name) in workbook.sheetnames:
        new_sheet = workbook.create_sheet(str(job_name))
    else:
        new_sheet = workbook[str(job_name)]
    
    #Casillas de información
    new_sheet['A1'] = str(job_name) +'-'+str(job_ID)
    new_sheet['A2'] = 'Names'
    
    #Agregar Nombres
    for i in range(len(name_list)):
        new_sheet['A'+str(i+3)] = name_list[i]
    
    #Agregar fechas
    for i in range(len(date_list)):
        new_sheet[str(chr(66+i))+'2'] = date_list[i]
        
    #Agregar horas
    job_day = datetime.strptime(job_day, "%m/%d/%y")
    job_day = (str(job_day.strftime('%A'))+" "+ str(job_day.day))
    
    for i in job_names:
        if i in name_list and job_day in date_list:
            row = name_list.index(str(i))
            column = date_list.index(str(job_day))
            if new_sheet[str(chr(66+int(column)))+str(int(row)+3)].value == None:
                new_sheet[str(chr(66+int(column)))+str(int(row)+3)] = '='+str("{:.2f}".format(hours[job_names.index(str(i))]))
            else:
                new_sheet[str(chr(66+int(column)))+str(int(row)+3)] += '+'+str("{:.2f}".format(hours[job_names.index(str(i))]))
        else:
            print('No coincide la fecha:', job_day, ', en el trabajo: ', job_name)
            pass
    
    #Llenar espacios en blanco de las horas con ceros
    for fila in range(3, int(len(name_list)+3)):
        for columna in range(2,9):
            if new_sheet.cell(row=fila, column=columna).value == None:
                new_sheet.cell(row=fila, column=columna, value='=0')
            else:
                pass
    
    #Agregar celdas con horas diarias
    cell_text_perday = ['TOTAL HOURS DAY - DAILY', 'TOTAL REGULAR HOURS - DAILY', 'TOTAL OVERTIME HOURS - DAILY']
    for i in range(len(cell_text_perday)):
        new_sheet['A'+str(len(name_list)+3+i)] = cell_text_perday[i]
        #Para este caso se tiene que editar pensando en que las celdas tienen valores dados otras tablas según la plantilla que se usa generalmente, es decir,
        #los valores de la casilla Total Regular Hours - Daily debe variar si hay overtime en alguna de las casillas
        for j in range(len(date_list)):
            if i==0:
                new_sheet[str(chr(66+j))+str(len(name_list)+3+i)] = "=SUM("+str(chr(66+j))+'3:'+str(chr(66+j))+str(len(name_list)+2)+')'
            elif i==1:
                new_sheet[str(chr(66+j))+str(len(name_list)+3+i)] = "="+str(chr(66+j))+str(len(name_list)+3)+"-"+str(chr(66+j))+str(len(name_list)+5)
            elif i==2:
                new_sheet[str(chr(66+j))+str(len(name_list)+3+i)] = "=0"
    
    #Formato y código de las horas semanales            
    cell_text_week = ['TOTAL HOURS - WEEKLY', 'TOTAL REGULAR HOURS - WEEKLY', 'TOTAL OVERTIME HOURS - WEEKLY']
    for i in range(len(cell_text_week)):
        new_sheet[str(chr(66+len(date_list)+i))+'2'] = cell_text_week[i]
        for j in range(len(name_list)):
            if i==0:
                new_sheet[str(chr(66+len(date_list)+i))+str(3+j)] = "=SUM("+str(chr(66))+str(3+j)+':'+str(chr(66+len(date_list)-1))+str(3+j)+')'
            elif i==1:
                new_sheet[str(chr(66+len(date_list)+i))+str(3+j)] = "=I"+str(3+j)+"-K"+str(3+j)
            elif i==2:
                new_sheet[str(chr(66+len(date_list)+i))+str(3+j)] = "=0"
    
    #Realiza la suma de cada una de las columnas
    new_sheet['I13'] = '=SUM(I3:I'+str(len(name_list)+2)+')'
    new_sheet['J14'] = '=SUM(J3:J'+str(len(name_list)+2)+')'
    new_sheet['K15'] = '=SUM(K3:K'+str(len(name_list)+2)+')'
    
    #Llena las celdas vacías
    for fila in range(int(len(name_list)+3),int(len(name_list)+6)):
        for columna in range(9,12):
            if new_sheet.cell(row=fila, column=columna).value == None:
                new_sheet.cell(row=fila, column=columna, value='-')
            else:
                pass
    
    workbook.save('test_savedata.xlsx')
    workbook.close()

def general_hours(name_list, date_list):
    workbook = load_workbook('test_savedata.xlsx')
    
    sheet_names = workbook.sheetnames
    sheet_names = [sheet_names[i+1] for i in range(len(sheet_names)-1)]
    sheet = workbook['General']
    #excel_general_sum = '='
    
    for i in range(len(name_list)):
        for j in range(len(date_list)):
            excel_general_sum = "=SUM('"+str(sheet_names[0])+":"+str(sheet_names[-1])+"'!"+str(chr(66 + int(j)))+str(int(i) + 3)+")"
            sheet[str(chr(66+int(j)))+str(int(i)+3)] = excel_general_sum
    workbook.save('test_savedata.xlsx')
    workbook.close()
    
if __name__ == "__main__":
    #Crea o verifica que las carpetas donde se almacenarán los archivos estén creadas
    folders = ['PDFs', 'SVGs', 'data','Excel']
    for i in folders:
        folder_creator(str(os.getcwd()),i)
    files, files_names = check_pdfs(os.path.join(str(os.getcwd()),'PDFs'))
    
    #Introducción de los días inicial y final de la semana de trabajo
    start_day = input('Fecha de inicio (year-month-day): ')
    start_day_fmt = datetime.strptime(start_day, '%Y-%m-%d')
    end_day = (start_day_fmt + timedelta(days=6)).strftime('%Y-%m-%d')
    #end_day = input('Fecha final (year-month-day): ')
    
    #Extrae los datos que se usan en todas las hojas de Excel, como los nombres.
    all_names = []
    day_list = daylist_generator(start_day,end_day)
    for i in range(len(files)):
        names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i])
        all_names = names + all_names
    all_names = [str(x) for x in np.unique(all_names)]
    all_names.sort()
    
    #Crea el archivo final de excel con la hoja 'General'
    excel_creator(all_names,day_list)
    
    #Agrega variables generales a las hojas nuevas de excel, y agrega horas
    for i in range(len(files)):
        names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i])
        new_sheets(all_names,job_name,job_ID,day_list,start_day_job,names,total_hours)
        start_day_job = datetime.strptime(start_day_job, "%m/%d/%y")
        start_day_job = (str(start_day_job.strftime('%A'))+" "+ str(start_day_job.day))

    #Agregar la suma total en la hoja general
    general_hours(all_names,day_list)
    
    #Texto comprobante
    print('우유')
    
    #for i in range(len(files)):
    #    pdf2svg(files[i],files_names[i])
    #for i in range(len(files)):
    #    pdf2excel(files[i],files_names[i])
    
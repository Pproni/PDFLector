from LectorPDF import *
#import LectorPDF
from win32com.client import Dispatch

def just_open(filename):
    xlApp = Dispatch("Excel.Application")
    current_directory = os.getcwd()
    file_path = os.path.join(current_directory, filename)
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(file_path)
    xlBook.Save()
    xlBook.Close()

def apply_cambios(names, job_name, cambios, cambios_jobs):
    #Cambio de los nombres a los nuevos de los empleados editados
    for i in range(len(cambios)):
        if cambios[i][0] in names:
            indice = names.index(cambios[i][0])
            names[indice] = cambios[i][1]
        else:
            pass
    #Cambio de los nombres a los nuevos de los trabajos editados
    for i in range(len(cambios_jobs)):
        if cambios_jobs[i][0] == job_name:
            job_name = cambios_jobs[i][1]
        else:
            pass
    return names, job_name

def Overtime_checker(name_list,date_list, overtime_data):
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
    
    just_open('test_savedata.xlsx')
    
    workbook = load_workbook('test_savedata.xlsx')
    workbook_values = load_workbook('test_savedata.xlsx', data_only=True)
    
    espacio_pagos = 1
    
    General_sheet = workbook['General']
    General_sheet_values = workbook_values['General']
    
    overtime_day_checker = 0
    #Esta parte revisa cuál es el primer día con overtime en la tabla de solo overtime
    for i in range(len(date_list)):
        if General_sheet[get_excel_column_name(29 + i + espacio_pagos) + str(len(name_list) + 3)].value != 0:
            overtime_day_checker = i
            break
        else:
            continue
    
    #Esta parte filtrará la lista con los días que tengan overtime
    for i in range(overtime_day_checker,len(date_list)):
        filter_list = []
        #Acá se seleccionan solo los datos que tengan el día que se está revisando el overtime
        filter_list = [x for x in overtime_data if x[0]==date_list[i]]
        #Se hace un bucle para revisar todas las personas que están en el trabajo
        for j in range(len(name_list)):
            filter_list_iterable = list(filter_list)
            #Se revisa que el valor del overtime en el día que se está revisando sea distinto de cero
            overtime_value_day = float(0)
            if (General_sheet_values[get_excel_column_name(29 + i + espacio_pagos) + str(j+ 3)].value) != 0:
                #Se extrae el overtime de la persona
                overtime_value_day = float(General_sheet_values[get_excel_column_name(29 + i + espacio_pagos) + str(j+ 3)].value)
                #Se filtra nuevamente la lista buscando solo los elementos que contengan el nombre de la persona que se está revisando
                filter_list_iterable = [sublista for sublista in filter_list_iterable if name_list[j] in sublista[2]]
                #Se crea una nueva lista que ordene cuál de las sublistas de filter_list_iterable termina más tarde,
                # es decir que su hora de trabajo terminado tiene un valor mayor
                filter_by_hour = []
                for indice, sublista in enumerate(filter_list_iterable):
                    indice_sublista = sublista[2].index(name_list[j])
                    filter_by_hour.append([indice,indice_sublista,sublista[3][indice_sublista],sublista[4][indice_sublista],sublista[1]])
                filter_by_hour.sort(key=lambda x: x[2], reverse=True)
                #print(name_list[j],date_list[i] ,filter_by_hour)
                for k in range(len(filter_by_hour)):
                    #Se crean las hojas que se van a trabajar, con fórmulas de excel y valores resultado de las fórmulas
                    Sheet_job = workbook[f'{filter_by_hour[k][-1]}']
                    Sheet_job_values = workbook_values[f'{filter_by_hour[k][-1]}']
                    if float(filter_by_hour[k][3]) >= float(overtime_value_day):
                        #Se actualiza el valor de cada Sheet, fórmula y valores, para que se sumen los resultados
                        ##Primero la hoja de fórmulas
                        Sheet_job[get_excel_column_name(i + 2) + str(5 + len(name_list))] = f'={float(Sheet_job_values[get_excel_column_name(i + 2) + str(5 + len(name_list))].value) + float(overtime_value_day)}'
                        Sheet_job['K'+ str(j + 3)] = f'={float(Sheet_job_values["K"+ str(j + 3)].value) + float(overtime_value_day)}'
                        ##Segundo la hoja de valores
                        Sheet_job_values[get_excel_column_name(i + 2) + str(5 + len(name_list))] = f'{float(Sheet_job_values[get_excel_column_name(i + 2) + str(5 + len(name_list))].value) + float(overtime_value_day)}'
                        Sheet_job_values['K'+ str(j + 3)] = f'{float(Sheet_job_values["K"+ str(j + 3)].value) + float(overtime_value_day)}'
                        break
                    elif float(filter_by_hour[k][3]) < float(overtime_value_day):
                        ##Primero la hoja de fórmulas
                        Sheet_job[get_excel_column_name(i + 2) + str(5 + len(name_list))] = f'={float(Sheet_job_values[get_excel_column_name(i + 2) + str(5 + len(name_list))].value) + float(filter_by_hour[k][3])}'
                        Sheet_job['K'+ str(j + 3)] = f'={float(Sheet_job_values["K"+ str(j + 3)].value) + float(filter_by_hour[k][3])}'
                        ##Segundo la hoja de valores
                        Sheet_job_values[get_excel_column_name(i + 2) + str(5 + len(name_list))] = f'{float(Sheet_job_values[get_excel_column_name(i + 2) + str(5 + len(name_list))].value) + float(filter_by_hour[k][3])}'
                        Sheet_job_values['K'+ str(j + 3)] = f'{float(Sheet_job_values["K"+ str(j + 3)].value) + float(filter_by_hour[k][3])}'
                        #Reducción
                        overtime_value_day = float(overtime_value_day) - float(filter_by_hour[k][3])
    workbook.save('test_savedata.xlsx')
    workbook_values.close()
    workbook.close()
    
def Supervisor_boxes(name_list, all_job_names):
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
    
    just_open('test_savedata.xlsx')
    
    workbook = load_workbook('test_savedata.xlsx')
    workbook_values = load_workbook('test_savedata.xlsx', data_only=True)
    General_sheet = workbook['General']
    
    sheet_names = workbook.sheetnames
    sheet_names.pop(0)
    
    thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))
    
    Supervisores = ['Andry Matos', 'Andry Matos']
    General_sheet[f'H{len(name_list) + 7}'].fill = PatternFill(fill_type="solid", fgColor="ea9999")
    General_sheet[f'I{len(name_list) + 7}'].fill = PatternFill(fill_type="solid", fgColor="ea9999")
    General_sheet[f'H{len(name_list) + 8}'] = Supervisores[0]
    General_sheet[f'H{len(name_list) + 8}'].fill = PatternFill(fill_type="solid", fgColor="ea9999")
    General_sheet[f'H{len(name_list) + 10}'] = Supervisores[1]
    General_sheet[f'H{len(name_list) + 10}'].fill = PatternFill(fill_type="solid", fgColor="ea9999")
    General_sheet.merge_cells(start_row=(len(name_list) + 8), start_column=8, end_row=(len(name_list) + 9), end_column=8)
    General_sheet.merge_cells(start_row=(len(name_list) + 10), start_column=8, end_row=(len(name_list) + 11), end_column=8)
    
    text_hours = ['Horas regulares', 'Horas overtime']
    for i in range(len(text_hours)):
        General_sheet[f'I{len(name_list) + 8 + 2*i}'] = text_hours[0]
        General_sheet[f'I{len(name_list) + 8 + 2*i}'].fill = PatternFill(fill_type="solid", fgColor="f9cb9c")
        General_sheet[f'I{len(name_list) + 9 + 2*i}'] = text_hours[1]
        General_sheet[f'I{len(name_list) + 9 + 2*i}'].fill = PatternFill(fill_type="solid", fgColor="a4c2f4")
    
    Pedro_index = name_list.index(Supervisores[0])
    Roberto_index = name_list.index(Supervisores[1])
    #if Supervisores[1] in name_list:
    #    Roberto_index = name_list.index(Supervisores[1])
    #else:
    #    Roberto_index = name_list.index(Supervisores[0])
    
    counter = 0
    for i in range(len(sheet_names)):
        sheet_job = workbook_values[f'{sheet_names[i]}']
        if float(sheet_job[f'J{Pedro_index + 3}'].value) != float(0) or float(sheet_job[f'K{Pedro_index + 3}'].value) != float(0):
            General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 7}'] = f'{sheet_names[i]}'
            General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 7}'].fill = PatternFill(fill_type="solid", fgColor="ea9999")
            for j in range(len(text_hours)):
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 8 + 2*j}'] = f"='{sheet_names[i]}'!J{name_list.index(Supervisores[j]) + 3}"
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 8 + 2*j}'].number_format = '0.00'
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 8 + 2*j}'].fill = PatternFill(fill_type="solid", fgColor="b7b7b7")
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 9 + 2*j}'] = f"='{sheet_names[i]}'!K{name_list.index(Supervisores[j]) + 3}"
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 9 + 2*j}'].number_format = '0.00'
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 9 + 2*j}'].fill = PatternFill(fill_type="solid", fgColor="b7b7b7")
            counter += 1
        elif float(sheet_job[f'J{Roberto_index + 3}'].value) != float(0) or float(sheet_job[f'K{Roberto_index + 3}'].value) != float(0):
            General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 7}'] = sheet_names[i]
            General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 7}'].fill = PatternFill(fill_type="solid", fgColor="ea9999")
            for j in range(len(text_hours)):
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 8 + 2*j}'] = f"='{sheet_names[i]}'!J{name_list.index(Supervisores[j]) + 3}"
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 8 + 2*j}'].number_format = '0.00'
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 8 + 2*j}'].fill = PatternFill(fill_type="solid", fgColor="b7b7b7")
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 9 + 2*j}'] = f"='{sheet_names[i]}'!K{name_list.index(Supervisores[j]) + 3}"
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 9 + 2*j}'].number_format = '0.00'
                General_sheet[f'{get_excel_column_name(10 + counter)}{len(name_list) + 9 + 2*j}'].fill = PatternFill(fill_type="solid", fgColor="b7b7b7")
            counter += 1
        else:
            pass
    
    centro_style = Alignment(horizontal="center", vertical="center",wrap_text=True)
    for c in range(8,counter + 10):
        for r in range(len(name_list) + 7,len(name_list) + 12):
            General_sheet.cell(row=r, column=c).alignment = centro_style
            General_sheet.cell(row=r, column=c).font = Font(name='Arial', size=10)
            General_sheet.cell(row=r, column=c).border = thin_border
    
    
    workbook.save('test_savedata.xlsx')
    workbook_values.close()
    workbook.close()
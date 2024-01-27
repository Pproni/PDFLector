from LectorPDF import *
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
            overtime_value_day = 0
            if (General_sheet_values[get_excel_column_name(29 + i + espacio_pagos) + str(j+ 3)].value) != 0:
                #Se extrae el overtime de la persona
                overtime_value_day = General_sheet_values[get_excel_column_name(29 + i + espacio_pagos) + str(j+ 3)].value
                #Se filtra nuevamente la lista buscando solo los elementos que contengan el nombre de la persona que se está revisando
                filter_list_iterable = [sublista for sublista in filter_list_iterable if name_list[j] in sublista[2]]
                #Se crea una nueva lista que ordene cuál de las sublistas de filter_list_iterable termina más tarde,
                # es decir que su hora de trabajo terminado tiene un valor mayor
                filter_by_hour = []
                for indice, sublista in enumerate(filter_list_iterable):
                    indice_sublista = sublista[2].index(name_list[j])
                    filter_by_hour.append([indice,indice_sublista,sublista[3][indice_sublista],sublista[4][indice_sublista],sublista[1]])
                filter_by_hour.sort(key=lambda x: x[2], reverse=True)
                print(name_list[j],date_list[i] ,filter_by_hour)
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
from LectorPDF import *
import Overtime
import SQLite_Database as SQLil
import Calendar
import difflib

files, files_names = check_pdfs(os.path.join(str(os.getcwd()),'PDFs'))
global place
place = input('Chicago o Celtic: ').lower()

def Conversion_process():
    # Crea o verifica que las carpetas donde se almacenarán los archivos estén creadas
    folders = ['PDFs', 'SVGs', 'data','Excel']
    for i in folders:
        folder_creator(str(os.getcwd()),i)

    files2, files_names2 = check_pdfs(os.path.join(str(os.getcwd()),'Excel'))
    
    s_files = set(files2)
    s_names = set(files_names2)
    difference_files = [x for x in files if x not in s_files]
    difference_names = [x for x in files_names if x not in s_names]
    
    if place == 'chicago':
        if difference_names:
            for i in range(len(difference_files)):
                pdf2svg(difference_files[i],difference_names[i])
            for i in range(len(difference_files)):
                pdf2excel(difference_files[i],difference_names[i],place)
        else:
            print('Ya está todo')
            pass
    elif place == 'celtic':
        if difference_names:
            for i in range(len(difference_files)):
                pdf2excel(difference_files[i],difference_names[i], place)
        else:
            print('Ya está todo')
            pass
                    

def Data_process():
    #Introducción de los días inicial y final de la semana de trabajo
    start_day = input('Fecha de inicio (year-month-day): ')
    #start_day = Calendar.mostrar_calendario()
    start_day_fmt = datetime.strptime(start_day, '%Y-%m-%d')
    end_day = (start_day_fmt + timedelta(days=6)).strftime('%Y-%m-%d')
    end_day_fmt = start_day_fmt + timedelta(days=6)
    #end_day = input('Fecha final (year-month-day): ')
    
    #Extrae los datos que se usan en todas las hojas de Excel, como los nombres.
    all_names = []
    all_job_names = []
    all_job_IDs = []
    hours_per_job = []
    day_list = daylist_generator(start_day,end_day)
    if place == 'chicago':
        for i in range(len(files)):
            names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i],place)
            hours_per_job.append(total_hours)
            all_names = names + all_names
            all_job_names.append(str(job_name))
            all_job_IDs.append(str(job_ID))
        all_names = [str(x) for x in np.unique(all_names)]
        all_names.sort()
    elif place == 'celtic':
        number_people = int(input('Número de empleados: '))
        for _ in range(number_people):
            all_names.append(input('Nombre a agregar: '))
        for i in range(len(files)):
            names, job_name, job_ID, total_hours,*_ = data_extractor(files_names[i], place)
            all_job_names.append(str(job_name))
            all_job_IDs.append(str(job_ID))
            hours_per_job.append(total_hours)
            #if names in all_names:
                
        
    #Revisión de base de datos
    cambios_DB = SQLil.getValuesList('Names_changes','Initial_Date',start_day)
    cambios_jobs_DB = SQLil.getValuesList('Jobs_changes','Initial_Date',start_day)    
        
    #Interfase de revisión
    all_names, all_job_names, all_job_IDs, cambios, cambios_jobs = Interface.interface(all_names,all_job_names,all_job_IDs, cambios_DB, cambios_jobs_DB, start_day)
    #print(all_names, 'check 1')
    #Reordenamiento de la lista de nombres
    all_names = [str(x) for x in np.unique(all_names)]
    all_names.sort()
    #print(all_names, 'check 2')
    
    #Crea el archivo final de excel con la hoja 'General'
    excel_creator(all_names,day_list,start_day_fmt,end_day_fmt)    
    start_days_perjob = []
    
    #Lista con los datos para el overtime
    Overtime_data = []
    
    #Agrega variables generales a las hojas nuevas de excel, y agrega horas
    for i in range(len(files)):
        #Extrae los datos
        names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i], place)
        #print(names,total_hours,job_name, 'check 3')
        if place == 'celtic':
            def compare_words_levenshtein_advanced(word1, word2, threshold=0.8):
                # Ignorar mayúsculas y espacios adicionales
                word1_clean = word1.lower().replace(" ", "")
                word2_clean = word2.lower().replace(" ", "")

                # Calcular el ratio de similitud
                ratio = difflib.SequenceMatcher(None, word1_clean, word2_clean).ratio()
                return ratio >= threshold

            # Aplicar la similitud de Levenshtein entre la lista names y all_names, quitando espacios y mayúsculas
            # y devolviendo el nombre en el formato de all_names, además de capturar los índices eliminados
            def filter_similar_names_with_removed_indices(names, all_names, threshold=0.8):
                result = []
                removed_indices = []  # Lista para almacenar los índices de los nombres que son eliminados

                for index, name in enumerate(names):
                    found_match = False
                    for ref_name in all_names:
                        if compare_words_levenshtein_advanced(name, ref_name, threshold):
                            result.append(ref_name)  # Devolver el nombre en el formato de all_names
                            found_match = True
                            break
                        
                    # Si no se encontró coincidencia, agregar el índice a la lista de eliminados
                    if not found_match:
                        removed_indices.append(index)           

                return result, removed_indices

            # Llamar a la función para filtrar los nombres y obtener los índices eliminados
            names, removed_indices = filter_similar_names_with_removed_indices(names, all_names)
            
            def remove_by_indices(original_list, removed_indices):
                return [item for idx, item in enumerate(original_list) if idx not in removed_indices]
            total_hours = remove_by_indices(total_hours, removed_indices)
            #names = filter_similar_names_advanced(names, all_names)
            #names = [x for x in names if x in all_names]
            #print(names, 'check 4')
        #Cambiar el formato de time_list para cuando el trabajo termine en el día siguiente al que comenzó
        for j in range(len(time_list)):
            if int(time_list[j][0]) > int(time_list[j][-1]):
                time_list[j][-1] = int(time_list[j][-1])+2400
        #Aplica los cambios hechos en la interfaz a las variables nombres de empleados y de trabajos
        names, job_name = Overtime.apply_cambios(names,job_name,cambios,cambios_jobs)
        #Crea las hojas del excel y acomoda las horas
        new_sheets(all_names,all_job_names[i],all_job_IDs[i],day_list,start_day_job,names,total_hours)
        #Cambia el formato del día de inicio del trabajo
        start_day_job = datetime.strptime(start_day_job, "%m/%d/%y")
        start_days_perjob.append(start_day_job.day)
        start_day_job = (str(start_day_job.strftime('%A'))+" "+ str(start_day_job.day))
        #Agrega los datos necesarios para el overtime en una lista
        Overtime_data.append([start_day_job,job_name,names,time_list, total_hours])
            
    #Cortafuegos
    Cortafuegos(all_names,all_job_names,hours_per_job, start_days_perjob,start_day_fmt)
    
    #Agregar la suma total en la hoja general
    general_hours(all_names,day_list)
    
    #Overtime
    Overtime.Overtime_checker(all_names,day_list,Overtime_data)
    
    #Tabla de supervisores
    Overtime.Supervisor_boxes(all_names, all_job_names)
    
    #Texto comprobante
    print('우유')
    
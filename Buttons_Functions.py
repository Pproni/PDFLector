from LectorPDF import *
import threading

files, files_names = check_pdfs(os.path.join(str(os.getcwd()),'PDFs'))

def Conversion_process():
    # Crea o verifica que las carpetas donde se almacenarán los archivos estén creadas
    folders = ['PDFs', 'SVGs', 'data','Excel']
    for i in folders:
        folder_creator(str(os.getcwd()),i)

    files2, files_names2 = check_pdfs(os.path.join(str(os.getcwd()),'data'))
    
    s_files = set(files2)
    s_names = set(files_names2)
    difference_files = [x for x in files if x not in s_files]
    difference_names = [x for x in files_names if x not in s_names]
   
    if difference_names:
        for i in range(len(difference_files)):
            pdf2svg(difference_files[i],difference_names[i])
        for i in range(len(difference_files)):
            pdf2excel(difference_files[i],difference_names[i])
    else:
        print('Ya está todo')
        pass
    
    #for i in range(len(files)):
    #    pdf2svg(files[i],files_names[i])
    #for i in range(len(files)):
    #    pdf2excel(files[i],files_names[i])

def Data_process():
    #Introducción de los días inicial y final de la semana de trabajo
    start_day = input('Fecha de inicio (year-month-day): ')
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
    for i in range(len(files)):
        names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i])
        hours_per_job.append(total_hours)
        all_names = names + all_names
        all_job_names.append(str(job_name))
        all_job_IDs.append(str(job_ID))
    all_names = [str(x) for x in np.unique(all_names)]
    all_names.sort()
    
    #Interfase de revisión
    all_names, all_job_names, all_job_IDs, cambios = Interface.interface(all_names,all_job_names,all_job_IDs)
    
    all_names = [str(x) for x in np.unique(all_names)]
    all_names.sort()
    
    #Crea el archivo final de excel con la hoja 'General'
    excel_creator(all_names,day_list,start_day_fmt,end_day_fmt)    
    start_days_perjob = []
    #Agrega variables generales a las hojas nuevas de excel, y agrega horas
    for i in range(len(files)):
        names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i])
        new_sheets(all_names,all_job_names[i],all_job_IDs[i],day_list,start_day_job,names,total_hours,cambios)
        start_day_job = datetime.strptime(start_day_job, "%m/%d/%y")
        start_days_perjob.append(start_day_job.day)
        start_day_job = (str(start_day_job.strftime('%A'))+" "+ str(start_day_job.day))
    
    #Cortafuegos
    Cortafuegos(all_names,all_job_names,hours_per_job, start_days_perjob,start_day_fmt)
    
    #Agregar la suma total en la hoja general
    general_hours(all_names,day_list)
    
    #Texto comprobante
    print('우유')
    
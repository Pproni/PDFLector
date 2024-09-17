from LectorPDF import *
import os
#import LectorPDF 

files, files_names = check_pdfs(os.path.join(str(os.getcwd()),'PDFs'))
place = 'celtic'

def ampm_24h(hora_am_pm):
        return datetime.strptime(hora_am_pm, "%I:%M %p").strftime("%H:%M")

#print(" ".isalpha())

#print(ampm_24h('5:45 PM'))
for i in range(len(files)):
    print('                         ')
    #print(files_names[i])
    names, job_name, job_ID, total_hours, time_list, start_day_job, end_day_job = data_extractor(files_names[i],place)
    print(start_day_job)
    
    print('ACABÓ')
    #print(names)



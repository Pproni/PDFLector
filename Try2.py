import os

def folder_creator(folder_directory,folder_name):
    path = os.path.join(folder_directory, folder_name)
    if not os.path.exists(path):
        os.mkdir(path)
        
folders = ['PDFs', 'SVGs', 'data','Excel']

for i in folders:
    folder_creator(str(os.getcwd()),i)
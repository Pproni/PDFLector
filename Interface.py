import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle
import SQLite_Database as SQLil

def interface(name_list,job_list,job_ID_list, cambios, cambios_jobs, initial_date):
    
    #cambios = []
    #cambios_jobs = []
    Old_name_list = list(name_list)
    Old_job_list = list(job_list)
    check_names = False
    check_jobs = False
    if cambios:
        for i in range(len(cambios)):
            if cambios[i][0] in name_list:
                check_names = True
                indice = name_list.index(cambios[i][0])
                name_list[indice] = cambios[i][1]
    if cambios_jobs:
        for i in range(len(cambios_jobs)):
            if cambios_jobs[i][0] in job_list:
                check_jobs = True
                indice = job_list.index(cambios_jobs[i][0])
                job_list[indice] = cambios_jobs[i][1]
        
    #Old_job_ID_list = list(job_ID_list)
    
    def cambiar_tema(tema):
        style.set_theme(tema)

    def on_select(event):
        selected_index = lista_nombres.focus()
        if selected_index:
            selected_values = lista_nombres.item(selected_index)['values']
            if selected_values and len(selected_values) > 0:
                selected_radio_option = opcion_var.get()
                selected_option = combobox.get()
                cuadro_texto_name_list.delete(0, tk.END)
                if selected_radio_option == "Opción 1" and selected_option == 'Job list':  # "Job name" seleccionado
                    cuadro_texto_name_list.insert(0, selected_values[1])
                elif selected_radio_option == "Opción 2" and selected_option == 'Job list':  # "Job ID" seleccionado
                    cuadro_texto_name_list.insert(0, selected_values[2])
                elif selected_option == 'Name list':
                    cuadro_texto_name_list.insert(0, selected_values[1])

    def actualizar_lista():
        lista_nombres.delete(*lista_nombres.get_children())
        for item in range(len(name_list)):
            lista_nombres.insert("", index='end', values=(Old_name_list[item], name_list[item]))

    def on_actualizar_lista():
        # Obtener el índice de la fila seleccionada
        selected_items = lista_nombres.selection()

        if selected_items:
            selected_index = lista_nombres.index(selected_items[0])
            selected_option = combobox.get()
            selected_radio_option = opcion_var.get()
            
            # Obtener el nuevo valor del cuadro de texto y el valor viejo
            nuevo_valor = cuadro_texto_name_list.get()

            # Actualizar el valor en la lista de datos
            if selected_option == 'Name list':
                valor_viejo = Old_name_list[selected_index]
                if check_names == False:
                    cambios.append([name_list[selected_index], nuevo_valor])
                else:
                    listaindex = [valor_viejo,name_list[selected_index]]
                    if listaindex[0] == listaindex[1]:
                        cambios.append([name_list[selected_index], nuevo_valor])
                        SQLil.inserRow('Names_changes',initial_date,[name_list[selected_index], nuevo_valor])
                    else:
                        index_cambio = cambios.index(listaindex)
                        cambios[index_cambio][1] = nuevo_valor
                        SQLil.updatefields('Names_changes',valor_viejo,nuevo_valor)
                name_list[selected_index] = nuevo_valor
                # Actualizar el valor en el treeview usando el índice obtenido
                lista_nombres.item(selected_items[0], values=(valor_viejo,nuevo_valor))
            elif selected_option == 'Job list':
                if selected_radio_option == "Opción 1":
                    valor_viejo = Old_job_list[selected_index]
                    if check_jobs == False:
                        cambios_jobs.append([job_list[selected_index],nuevo_valor])
                    else:
                        listaindex = [valor_viejo,job_list[selected_index]]
                        if listaindex[0] == listaindex[1]:
                            cambios_jobs.append([job_list[selected_index],nuevo_valor])
                            SQLil.inserRow('Jobs_changes',initial_date,[job_list[selected_index],nuevo_valor])
                        else:
                            index_cambio = cambios_jobs.index(listaindex)
                            cambios_jobs[index_cambio][1] = nuevo_valor
                            SQLil.updatefields('Jobs_changes',valor_viejo,nuevo_valor)
                    job_list[selected_index] = nuevo_valor
                    lista_nombres.item(selected_items[0], values=(valor_viejo,nuevo_valor,job_ID_list[selected_index]))
                elif selected_radio_option == "Opción 2":
                    valor_viejo = Old_job_list[selected_index]
                    job_ID_list[selected_index] = nuevo_valor
                    lista_nombres.item(selected_items[0], values=(valor_viejo,job_list[selected_index], nuevo_valor))

            # Actualizar el valor en el treeview usando el índice obtenido
            #lista_nombres.item(selected_items[0], values=(nuevo_valor,))

            # Limpiar el cuadro de texto
            cuadro_texto_name_list.delete(0, tk.END)
        else:
            # Mostrar un mensaje o manejar de alguna manera que no hay elementos seleccionados
            print("No hay elementos seleccionados para actualizar.")


    def on_radio_select():
        selected_index = lista_nombres.focus()
        
        if selected_index:
            selected_values = lista_nombres.item(selected_index)['values']
            if selected_values and len(selected_values) > 1:
                selected_option = opcion_var.get()
                cuadro_texto_name_list.delete(0, tk.END)
                
                if selected_option == "Opción 1":  # "Job name" seleccionado
                    cuadro_texto_name_list.insert(0, selected_values[1])
                elif selected_option == "Opción 2":  # "Job ID" seleccionado
                    cuadro_texto_name_list.insert(0, selected_values[2])
            else:
                print("El elemento seleccionado no tiene valores asociados.")
        else:
            print("Ningún elemento seleccionado.")

    def agregar_elemento():
        nuevo_elemento = cuadro_texto_name_list.get()
        if nuevo_elemento:
            name_list.append(nuevo_elemento)
            Old_name_list.append(nuevo_elemento)
            actualizar_lista()
            cuadro_texto_name_list.delete(0, tk.END)

    def mostrar_name_list():
        # Configurar el Treeview para mostrar 'Name list'
        lista_nombres["columns"] = ('Initial_Name','Final_Name')
        lista_nombres.heading('#1', text='Initial Name')
        lista_nombres.column('#1', anchor='center', stretch='yes')
        lista_nombres.heading('#2', text='Final Name')
        lista_nombres.column('#2', anchor='center', stretch='yes')
        actualizar_lista()

        # Mostrar los elementos asociados con "Name list"
        texto_seleccionado_name_list.grid(row=0, column=1, padx=0, pady=20, sticky="w")
        cuadro_texto_name_list.grid(row=0, column=2, padx=25, pady=20, sticky="w")
        boton_actualizar_name_list.grid(row=0, column=4, padx=20, pady=20, sticky="w")
        boton_agregar_name_list.grid(row=0, column=3, padx=20, pady=20, sticky="w")

        # Ocultar los botones de la opción "Job list"
        boton_opcion1.grid_forget()
        boton_opcion2.grid_forget()

    def mostrar_job_list():
        # Configurar el Treeview para mostrar 'Job list'
        lista_nombres["columns"] = ('Initial_JobName','Final_JobName' ,'Job ID list')
        lista_nombres.heading('#1', text='Initial Job Name')
        lista_nombres.heading('#2', text='Final Job Name')
        lista_nombres.heading('#3', text='Job ID list')
        lista_nombres.column('#1', anchor='center', stretch='yes')
        lista_nombres.column('#2', anchor='center', stretch='yes')
        lista_nombres.column('#3', anchor='center', stretch='yes')
        
        # Actualizar la lista con los elementos de job_list y job_ID_list
        lista_nombres.delete(*lista_nombres.get_children())
        for item1, item2 in zip(job_list, job_ID_list):
            index = job_list.index(item1)
            lista_nombres.insert("", index='end', values=(Old_job_list[index],item1, item2))
        
        # Mostrar los botones de la opción "Job list"
        boton_opcion1.grid(row=0, column=3, padx=39, sticky="w")
        boton_opcion2.grid(row=1, column=3, padx=39, sticky="w")

        # Ocultar los elementos asociados con "Name list"
        boton_agregar_name_list.grid_forget()


    def on_combobox_select(event):
        selected_option = combobox.get()

        # Mostrar elementos según la opción seleccionada
        if selected_option == 'Name list':
            opcion_var.set(value="Opción 1")
            cuadro_texto_name_list.delete(0, tk.END)
            mostrar_name_list()
            
        elif selected_option == 'Job list':
            opcion_var.set(value="Opción 1")
            cuadro_texto_name_list.delete(0, tk.END)
            mostrar_job_list()
        

    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.geometry('800x400')
    ventana.title("Editor de datos")

    # Crear un estilo temático con ttkthemes
    style = ThemedStyle(ventana)
    cambiar_tema('breeze')

    # Crear un widget Treeview con ttk
    lista_nombres = ttk.Treeview(ventana, columns=(), show='headings')
    lista_nombres.grid(row=2, column=0, padx=20, pady=20, sticky="nsew", columnspan=5)

    # Configurar un evento de selección
    lista_nombres.bind("<<TreeviewSelect>>", on_select)

    # Configurar el redimensionamiento de columnas y filas
    ventana.columnconfigure(0, weight=1)
    ventana.rowconfigure(1, weight=1)

    # Inicializar la lista en la Treeview
    actualizar_lista()

    # Crear una lista desplegable (Combobox) en la columna cero
    opciones_combobox = ['Name list', 'Job list']
    combobox = ttk.Combobox(ventana, values=opciones_combobox)
    combobox.set(opciones_combobox[0])
    combobox.grid(row=0, column=0, padx=25, pady=20, sticky="w")

    # Mostrar elementos iniciales
    texto_seleccionado_name_list = ttk.Label(ventana, text="Texto seleccionado:")
    cuadro_texto_name_list = tk.Entry(ventana, width=25)
    boton_agregar_name_list = ttk.Button(ventana, text="Agregar Elemento", command=agregar_elemento)
    boton_actualizar_name_list = ttk.Button(ventana, text="Actualizar Lista", command=on_actualizar_lista)

    # Configurar botones asociados con "Job list"
    opcion_var = tk.StringVar(value="Opción 1")
    boton_opcion1 = ttk.Radiobutton(ventana, text="Job name", variable=opcion_var, value="Opción 1", command=on_radio_select)
    boton_opcion2 = ttk.Radiobutton(ventana, text="Job ID", variable=opcion_var, value="Opción 2", command=on_radio_select)
    boton_opcion1.grid(row=0, column=3)
    boton_opcion2.grid(row=1, column=3)

    mostrar_name_list()

    # Configurar un evento de selección para la lista desplegable
    combobox.bind("<<ComboboxSelected>>", on_combobox_select)
    # Correr la aplicación
    ventana.mainloop()
    
    if check_names == False:
        SQLil.insertRows('Names_changes',initial_date, cambios)
        print('Se guardaron los cambios en los nombres de los empleados de esta nueva factura.')
    if check_jobs == False:
        SQLil.insertRows('Jobs_changes',initial_date, cambios_jobs)
        print('Se guardaron los cambios en los nombres de los trabajos de esta nueva factura.')
        
    return name_list,job_list,job_ID_list,cambios,cambios_jobs

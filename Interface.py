import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle

def interface(name_list,job_list,job_ID_list):
    
    cambios = []
    
    def cambiar_tema(tema):
        style.set_theme(tema)

    def on_select(event):
        selected_index = lista_nombres.focus()
        if selected_index:
            selected_values = lista_nombres.item(selected_index)['values']
            if selected_values and len(selected_values) > 0:
                cuadro_texto_name_list.delete(0, tk.END)
                cuadro_texto_name_list.insert(0, selected_values[0])

    def actualizar_lista():
        lista_nombres.delete(*lista_nombres.get_children())
        for item in name_list:
            lista_nombres.insert("", index='end', values=(item, f"Descripción {item}"))

    def obtener_indice_seleccionado():
        selected_index = lista_nombres.selection()[0]
        return lista_nombres.index(selected_index)

    def on_actualizar_lista():
        # Obtener el índice de la fila seleccionada
        selected_items = lista_nombres.selection()

        if selected_items:
            selected_index = lista_nombres.index(selected_items[0])
            selected_option = combobox.get()
            selected_radio_option = opcion_var.get()
            
            # Obtener el nuevo valor del cuadro de texto
            nuevo_valor = cuadro_texto_name_list.get()

            # Actualizar el valor en la lista de datos
            if selected_option == 'Name list':
                cambios.append([name_list[selected_index], nuevo_valor])
                name_list[selected_index] = nuevo_valor
                # Actualizar el valor en el treeview usando el índice obtenido
                lista_nombres.item(selected_items[0], values=(nuevo_valor,))
            elif selected_option == 'Job list':
                if selected_radio_option == "Opción 1":
                    job_list[selected_index] = nuevo_valor
                    lista_nombres.item(selected_items[0], values=(nuevo_valor,job_ID_list[selected_index]))
                elif selected_radio_option == "Opción 2":
                    job_ID_list[selected_index] = nuevo_valor
                    lista_nombres.item(selected_items[0], values=(job_list[selected_index],nuevo_valor))

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
                    cuadro_texto_name_list.insert(0, selected_values[0])
                elif selected_option == "Opción 2":  # "Job ID" seleccionado
                    cuadro_texto_name_list.insert(0, selected_values[1])
            else:
                print("El elemento seleccionado no tiene valores asociados.")
        else:
            print("Ningún elemento seleccionado.")

    def agregar_elemento():
        nuevo_elemento = cuadro_texto_name_list.get()
        if nuevo_elemento:
            name_list.append(nuevo_elemento)
            actualizar_lista()
            cuadro_texto_name_list.delete(0, tk.END)

    def mostrar_name_list():
        # Configurar el Treeview para mostrar 'Name list'
        lista_nombres["columns"] = ('Nombres',)
        lista_nombres.heading('#1', text='Nombres')
        lista_nombres.column('#1', anchor='center', stretch='yes')
        actualizar_lista()

        # Ocultar el botón de agregar
        # boton_agregar_name_list.grid_forget()

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
        lista_nombres["columns"] = ('Job list', 'Job ID list')
        lista_nombres.heading('#1', text='Job list')
        lista_nombres.heading('#2', text='Job ID list')
        lista_nombres.column('#1', anchor='center', stretch='yes')
        lista_nombres.column('#2', anchor='center', stretch='yes')
        
        # Actualizar la lista con los elementos de job_list y job_ID_list
        lista_nombres.delete(*lista_nombres.get_children())
        for item1, item2 in zip(job_list, job_ID_list):
            lista_nombres.insert("", index='end', values=(item1, item2))
        
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
    return name_list,job_list,job_ID_list,cambios

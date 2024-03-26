import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from ttkthemes import ThemedStyle
from datetime import datetime, timedelta

def mostrar_calendario():
    def cambiar_tema(tema):
        style.set_theme(tema)
    
    def click_calendario(event):
        eventos = cal.get_calevents(tag='Evento resaltado')
        if eventos:
            cal.calevent_remove()

        fecha_seleccionada = cal.get_date()
        fecha_seleccionada_obj = datetime.strptime(fecha_seleccionada, "%d/%m/%Y")
        fecha_seleccionada_formateada = fecha_seleccionada_obj.strftime("%Y-%m-%d")
        
        etiqueta_fecha.config(text=f"Fecha seleccionada: {fecha_seleccionada_formateada}", font="Arial 12")
        
        # Actualizar las variables globales
        global fecha_seleccionada_global
        fecha_seleccionada_global = str(fecha_seleccionada_formateada)

        for i in range(7):
            fecha_resaltar = fecha_seleccionada_obj + timedelta(days=i)
            cal.calevent_create(fecha_resaltar, 'resaltado', 'Evento resaltado')
            cal.tag_config("Evento resaltado", background='blue', foreground='white')

    def seleccionar_fecha():
        # Mostrar la fecha seleccionada global        
        if fecha_seleccionada_global:
            print("Fecha seleccionada:", fecha_seleccionada_global)
        ventana_calendario.destroy()

    ventana_calendario = tk.Tk()
    ventana_calendario.title("Calendario")
    
    # Crear un estilo temático con ttkthemes
    style = ThemedStyle(ventana_calendario)
    cambiar_tema('breeze')

    etiqueta_semana_trabajo = tk.Label(ventana_calendario, text="Escoja la semana de trabajo ha procesar:", font=("Arial", 12, "bold"))
    etiqueta_semana_trabajo.pack(pady=(17,1))

    cal = Calendar(ventana_calendario, selectmode="day", date_pattern="dd/MM/yyyy", weekenddays=[6, 7],
                   font="Arial 15", spacing=15)
    cal.pack(pady=15)

    cal.bind("<<CalendarSelected>>", click_calendario)

    etiqueta_fecha = ttk.Label(ventana_calendario, text="Fecha seleccionada: ", font="Arial 12")
    etiqueta_fecha.pack(pady=10)

    boton_hora = ttk.Button(ventana_calendario, text="Seleccionar fecha", command=seleccionar_fecha)
    boton_hora.pack(pady=5)

    etiqueta_hora = ttk.Label(ventana_calendario, text="", font="Arial 12")
    etiqueta_hora.pack(pady=10)

    # Variables globales para almacenar la fecha seleccionada
    global fecha_seleccionada_global
    fecha_seleccionada_global = ""

    ventana_calendario.mainloop()  # Mantener la ventana abierta
    return fecha_seleccionada_global

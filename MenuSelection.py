import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle
from Buttons_Functions import *
import threading

def cambiar_tema(tema):
    style.set_theme(tema)
    
def iniciar_barra_progreso():
    progress_window = tk.Toplevel()
    progress_window.title("Procesando...")
    progress_window.geometry('300x100')

    label = tk.Label(progress_window, text="Procesando, por favor espere...")
    label.pack(pady=20)

    progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
    progress_bar.pack()

    # Inicia la barra de progreso indeterminada
    progress_bar.start()

    return progress_window, progress_bar

def finalizar_barra_progreso(progress_window):
    progress_window.withdraw()  # Oculta la ventana de progreso
    progress_window.update()   # Actualiza la interfaz gráfica
    progress_window.destroy()   # Destruye la ventana de progreso

def ejecutar_conversion(progress_window):
    try:
        # Coloca aquí tu lógica real de conversión (sin time.sleep)
        Conversion_process()
    finally:
        # Detiene la barra de progreso y cierra la ventana de progreso
        ventana.deiconify()  # Vuelve a mostrar la ventana principal
        finalizar_barra_progreso(progress_window)

def mover_barra_progreso(progress_bar, progress_window):
    if not progress_window.winfo_exists():  # Verifica si la ventana principal aún existe
        return  # Detiene la actualización si la ventana principal ha sido cerrada

    try:
        # Incrementa el valor de la barra de progreso
        progress_bar.step(10)
    except tk.TclError:
        return  # Detiene la actualización si la barra de progreso ha sido destruida

    # Programa el próximo incremento después de un tiempo muy corto
    progress_bar.after(50, lambda: mover_barra_progreso(progress_bar, progress_window))

def seleccionar_conversion():
    print("Seleccionaste Conversión")
    # Cerrar la ventana actual con los botones
    ventana.withdraw()  # Oculta la ventana principal

    # Crea y muestra la ventana< de progreso
    progress_window, progress_bar = iniciar_barra_progreso()

    # Inicia la función mover_barra_progreso para aumentar la velocidad
    mover_barra_progreso(progress_bar, progress_window)

    # Ejecuta la función Conversion_process() en un hilo separado
    thread = threading.Thread(target=ejecutar_conversion, args=(progress_window,), name='Progress_bar')
    thread.start()

def seleccionar_organizacion():
    ventana.destroy()  # Cerrar la ventana actual
    Data_process()
    print("Seleccionaste Organización")
    abrir_ventana()  # Abrir una nueva ventana después del proceso

def abrir_ventana():
    # Variables globales
    global ventana
    global style

    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.geometry('300x150')
    ventana.title("Menú de Selección")

    # Crear un estilo temático con ttkthemes
    style = ThemedStyle(ventana)
    cambiar_tema('breeze')

    # Crear los botones de selección
    boton_conversion = ttk.Button(ventana, text="Conversión", command=seleccionar_conversion)
    boton_organizacion = ttk.Button(ventana, text="Organización", command=seleccionar_organizacion)

    # Centrar los botones en la ventana principal
    boton_conversion.pack(side=tk.TOP, pady=20)
    boton_organizacion.pack(side=tk.TOP, pady=20)

    # Iniciar el bucle de eventos de la ventana principal
    ventana.mainloop()

abrir_ventana()

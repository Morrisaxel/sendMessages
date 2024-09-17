#Script 17-08-2024

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit as pwk
import time
import pyautogui
import keyboard as k

# Crear ventana principal
root = tk.Tk()
root.title("Envío de Mensajes por WhatsApp")

# Variables globales
excel_data = None

# Función para validar el número de teléfono
def validar_telefono(telefono: str) -> bool:
    return telefono.isdigit() and len(telefono) == 9

# Función para cargar el archivo de Excel
def cargar_archivo():

    global excel_data
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            # Especificamos el motor según el tipo de archivo
            if file_path.endswith(".xlsx"):
                excel_data = pd.read_excel(file_path, engine='openpyxl')
            elif file_path.endswith(".xls"):
                excel_data = pd.read_excel(file_path, engine='xlrd')
            else:
                raise ValueError("Formato de archivo no compatible.")
            
            # Comprobamos que las columnas necesarias existan
            print(excel_data.columns)
            required_columns = ['CLIENTE', 'TELÉFONO', 'LINEA OFERTADA 1RA/ MONTO TOTAL LINEA 2DA/ RAPIDITO']
            if all(col in excel_data.columns for col in required_columns):
                # Filtramos solo las columnas necesarias y las renombramos
                excel_data = excel_data[required_columns]
                excel_data.columns = ['cliente', 'telefono', 'oferta']
                print(excel_data.head)
                messagebox.showinfo("Éxito", "Archivo cargado y columnas limpiadas correctamente")
            else:
                messagebox.showerror("Error", "El archivo debe tener las columnas 'CLIENTE', 'TELÉFONO' y 'LINEA OFERTADA 1RA'")
        except Exception as e:
            messagebox.showerror("Error", f"Hubo un problema al leer el archivo: {str(e)}")


# Función para enviar mensajes de WhatsApp
def enviar_mensajes():
    global excel_data
    if excel_data is None:
        messagebox.showerror("Error", "Primero debes cargar un archivo de Excel.")
        return
    
    for index, row in excel_data.iterrows():
        telefono = str(row['telefono'])
        print(telefono)
        cliente = row['cliente']
        oferta = row['oferta']
        mensaje = f"Hola buenas tardes {cliente}, le comunicamos que tiene una {oferta}."

        # Validar que el número sea correcto
        if validar_telefono(telefono):
            try:
                # # # # Enviar el mensaje por WhatsApp
                pwk.sendwhatmsg_instantly(f"+51{telefono[:9]}", message=mensaje, wait_time=10, tab_close =True, close_time=3)
                time.sleep(2)
                pyautogui.click()
                time.sleep(1)
                k.press_and_release('enter')
              
               

            except Exception as e:
                print(f"Error al enviar el mensaje a {telefono}: {str(e)}")
        else:
            print(f"Número de teléfono inválido: {telefono}")
    
    messagebox.showinfo("Éxito", "Todos los mensajes han sido enviados.")

# Configuración del GUI
# Input para el número de teléfono
label_telefono = tk.Label(root, text="Número de Teléfono (9 dígitos):")
label_telefono.pack(pady=5)
entry_telefono = tk.Entry(root)
entry_telefono.pack(pady=5)


# Botón para cargar el archivo de Excel
btn_cargar = tk.Button(root, text="Cargar Archivo de Excel", command=cargar_archivo)
btn_cargar.pack(pady=10)

# Botón para enviar mensajes
btn_enviar = tk.Button(root, text="Enviar Mensajes por WhatsApp", command=enviar_mensajes)
btn_enviar.pack(pady=10)

# Iniciar la ventana
root.mainloop()

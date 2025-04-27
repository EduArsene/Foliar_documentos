'''import ctypes
import os

# Ajustar el tamaño de la consola
os.system('mode con: cols=-20 lines=-25')

# Obtener el tamaño de la pantalla
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)

# Definir la posición de la consola (esquina inferior izquierda)
console_x = -20
console_y = screen_height - 50  # Altura mínima de la consola

# Mover la consola a la posición definida
hwnd = ctypes.windll.kernel32.GetConsoleWindow()
ctypes.windll.user32.MoveWindow(hwnd, console_x, console_y, 200, 50, True)  # 200x50 es el tamaño de la consola
'''
#
'''import ctypes
import tkinter as tk
#from pystray import Icon, MenuItem, Menu
from PIL import Image

# Función para ocultar la consola
def hide_console():
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd != 0:
        ctypes.windll.user32.ShowWindow(hwnd, 0)  # Ocultar la consola

# Función para mostrar la consola (opcional)
def show_console():
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd != 0:
        ctypes.windll.user32.ShowWindow(hwnd, 5)  # Mostrar la consola

# Función para salir de la aplicación desde la bandeja del sistema
def exit_app(icon, item):
    icon.stop()
    root.quit()

# Ocultar la consola al inicio
hide_console()

# Crear la ventana principal de Tkinter
root = tk.Tk()
root.geometry("400x300")
root.title("Aplicación Principal")

# Ejecutar la aplicación
root.mainloop()'''
'''
import tkinter as tk
from PIL import Image, ImageTk

# Crear la ventana principal
root = tk.Tk()
root.geometry("400x300")
root.title("Aplicación Principal")

# Cargar la imagen de fondo
bg_image = Image.open("Sunat.jpg")
bg_image = bg_image.resize((400, 300))  # Asegúrate de que la imagen tenga el tamaño de la ventana
bg_photo = ImageTk.PhotoImage(bg_image)

# Crear un Canvas y colocar la imagen de fondo
canvas = tk.Canvas(root, width=400, height=300)
canvas.pack(fill="both", expand=True)
canvas.create_image(0, 0, anchor="nw", image=bg_photo)

# Añadir widgets sobre el fondo
#label = tk.Label(root, text="¡Hola, mundo!", bg="#ffffff", font=("Arial", 16))
#canvas.create_window(200, 150, window=label)  # Coloca el widget en el centro del canvas

# Ejecutar la aplicación
root.mainloop()
'''

    #print(font)

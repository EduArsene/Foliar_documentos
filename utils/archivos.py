# --- utils/archivos.py ---
import os
from tkinter import messagebox

def hide_console():
    import ctypes
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd != 0:
        ctypes.windll.user32.ShowWindow(hwnd, 0)

def abrir_directorio(archivo):
    directorio = os.path.dirname(archivo)
    if os.path.exists(directorio):
        os.startfile(directorio)
    else:
        messagebox.showerror("Error", "No se pudo abrir el directorio.")

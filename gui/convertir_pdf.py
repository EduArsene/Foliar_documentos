#Ventana para convertir docx a PDF
import os
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

color_secundario = "#73A5BC"
def crear_interfaz():
    from docx2pdf import convert
    #cerrar la ventana principal
    root.iconify()
    #funcion para seleccionar el documento
    def seleccionar_archivo():
        archivo = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx"), ("Word Document", "*.DOCX")])
        if archivo:
            #cambiando de extension, ya que la funcion "convert" no acepta documentos con extensiones mayusculas
            new_file = archivo.replace(".DOCX",".docx")
            try:
                #reemplazamos el archivo con uno nuevo (copia) para que finalmente se guarde en minuscula al extension
                os.replace(archivo, new_file)
                ruta_archivo.set(new_file)
                ventana_convertir.update_idletasks()
            except OSError as e:
                messagebox.showerror(f"Error al renombrar el archivo: {e}")
                      
    #funcion para convertir los  docx
    def convertir_archivo():
        #pedimos la ruta del archivo en una variable
        input = ruta_archivo.get()
        if not input:
            messagebox.showerror("Error", "Seleccione un archivo .docx")
        else:    
            #crea un pdf
            archivo_pdf = os.path.splitext(input)[0] + ".pdf"
            estado.set("Convirtiendo...")
            ventana_convertir.update_idletasks()
            #convierte el archivo a PDF
            convert(input, archivo_pdf,True)
            #mostramos el nombre con el que se guardó el archivo
            estado.set(f"Finalizado: {os.path.basename(archivo_pdf)}")
            messagebox.showinfo("Éxito", f"El archivo ha sido convertido y guardado como: \n{os.path.basename(archivo_pdf)}")
    #minimizar la ventana principal            
    def minimizar():
        root.deiconify()
        ventana_convertir.destroy()
    # Crear la ventana principal
    root = tk.Tk()
    ventana_convertir = tk.Toplevel(root)
    ventana_convertir.title("Conversor")
    ventana_convertir.geometry("500x250")
    ventana_convertir.minsize(500,250)
    ventana_convertir.maxsize(500,250)
    ventana_convertir.configure(bg=color_secundario)  # Fondo azul claro  
    ventana_convertir.protocol("WM_DELETE_WINDOW",minimizar)
    ruta_archivo = tk.StringVar()
    estado = tk.StringVar()
    # Estilo para los widgets
    estilo = ttk.Style()
    estilo.theme_use('clam')
    estilo.configure('TLabel', background=color_secundario, font=('Arial', 10))
    estilo.configure('TEntry', font=('Arial', 10))
    estilo.configure('TButton', font=('Arial', 10))

    # Título
    titulo = ttk.Label(ventana_convertir, text="Convertir de DOCX a PDF", font=('Arial', 16, 'bold'), background=color_secundario)
    titulo.pack(pady=10)

    # Label e Entry
    label_info = ttk.Label(ventana_convertir, text="Ruta del archivo:")
    label_info.pack(pady=5)

    entry_info = ttk.Entry(ventana_convertir, width=60, state='readonly', textvariable=ruta_archivo)
    entry_info.pack()
    entry_info.insert(0, "Datos de aduana aquí")

    # Frame para los botones
    frame_examinar = ttk.Frame(ventana_convertir)
    frame_examinar.pack(pady=10)

    # Botones
    boton_procesar = ttk.Button(frame_examinar, text="Examinar", style='azul.TButton',command=seleccionar_archivo)
    boton_procesar.pack(side=tk.LEFT, padx=0)

    frame_convertir = ttk.Frame(ventana_convertir)
    frame_convertir.pack(pady=15)
    boton_convertir = ttk.Button(frame_convertir, text="Convertir", style='Rojo.TButton',command=convertir_archivo)
    boton_convertir.pack(side=tk.LEFT, padx=0)

    # Configurar estilos de los botones
    estilo.configure('azul.TButton', background='#096CB7', foreground='white')
    estilo.configure('Rojo.TButton', background='#BD0C4E', foreground='white')
    estilo.map('azul.TButton', background=[('active', '#005596')])
    estilo.map('Rojo.TButton', background=[('active', '#8B0034')])

    # Label de estado
    tk.Label(ventana_convertir, textvariable=estado,bg=color_secundario).pack(pady=10)

    ventana_convertir.mainloop()
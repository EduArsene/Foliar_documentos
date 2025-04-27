'cx-Oracle','cx_Freeze','teradata','auto-py-to-exe','pandas','ipykernel','jupyter_client','jupyter_core','matplotlib-inline','numpy','pillow','ttkbootstrap'

#para poder agreagar iconos al ejecutable
pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module matplotlib-inline --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onefile --icon=icono.ico Foliar.py
pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module matplotlib-inline --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onedir --icon=icono.ico Foliar.py

#El mejor para poder convertir en ejecutable
#en caso tengas mas librerias deberas aumentarlos de la siguiente manera" --exclude-module (nombreDeLibreria)
pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onefile --icon=icono.ico --add-data "src;src" Foliar.py
pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module matplotlib-inline --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onedir --icon=icono.ico --add-data "src;src" Foliar.py

#en caso quieras tener un ejecutador con una api para poder realizar diseños con html
pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module matplotlib-inline --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onedir --icon=icono.ico --add-data "src;src" prueba.py

pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module matplotlib-inline --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onedir --icon=icono.ico --add-data "D:\prueba2\html/index.html;." --add-data "D:\prueba2\html\\style.css;." "D:\prueba2\html\\style2.css;." "D:\prueba2\html\\segunda_interfaz.html;." --add-data "src;src" prueba.py


#
pyinstaller --exclude-module cx-Oracle --exclude-module cx_Freeze --exclude-module teradata --exclude-module auto-py-to-exe --exclude-module pandas --exclude-module ipykernel --exclude-module jupyter_client --exclude-module jupyter_core --exclude-module matplotlib-inline --exclude-module ttkbootstrap --exclude-module numpy --exclude-module pillow --onedir --icon=icono.ico --add-data "D:\prueba2\html\index.html;." --add-data "D:\prueba2\html\style.css;." --add-data "D:\prueba2\html\style2.css;." --add-data "D:\prueba2\html\segunda_interfaz.html;." --add-data "src;src" prueba.py
import importlib.resources as pkg_resources
import src

def abrir_imagen(nombre_archivo):
    with pkg_resources.path(src, nombre_archivo) as ruta_imagen:
        return ruta_imagen
# --- utils/docx_tools.py ---
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def agregar_elemento(agregar, tag, tipo, texto=None):
    elemento = OxmlElement(tag)
    elemento.set(qn('w:fldCharType'), tipo)
    if texto:
        elemento.set(qn('xml:space'), 'preserve')
        elemento.text = texto
    agregar._r.append(elemento)

def configurar_idioma_parrafo(parrafo, lang_code):
    for run in parrafo.runs:
        rPr = run._element.get_or_add_rPr()
        lang = OxmlElement('w:lang')
        lang.set(qn('w:val'), lang_code)
        rPr.append(lang)

def activar_pares_impares_diferentes(seccion):
    seccPr = seccion._sectPr
    if seccPr is None:
        seccPr = OxmlElement('w:sectPr')
        seccion._element.append(seccPr)
    for el in seccPr.findall(qn('w:titlePg')):
        seccPr.remove(el)
    evenAndOddHeaders = OxmlElement('w:evenAndOddHeaders')
    seccPr.append(evenAndOddHeaders)
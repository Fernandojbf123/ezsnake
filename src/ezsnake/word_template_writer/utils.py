"""
General utility functions for Word document manipulation.

This module contains generic helper functions that are useful across different
Word manipulation tasks.

Public utilities - Can be imported directly from the main package.
"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def insert_line_feed(paragraph, texto="", centrado=False):
    """Inserta un nuevo párrafo después del párrafo dado usando manipulación XML.
    
    Args:
        paragraph: El párrafo después del cual insertar.
        texto: El texto del nuevo párrafo (opcional).
        centrado: Si True, centra el texto del párrafo.
    
    Returns:
        El elemento XML del nuevo párrafo creado.
    """
    # Obtener el elemento del párrafo actual
    p_element = paragraph._element
    # Obtener el elemento padre
    parent = p_element.getparent()
    # Crear un nuevo elemento de párrafo
    nuevo_p = OxmlElement('w:p')
    
    # Si necesita estar centrado
    if centrado:
        pPr = OxmlElement('w:pPr')
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center')
        pPr.append(jc)
        nuevo_p.append(pPr)
    
    # Si hay texto, agregarlo
    if texto:
        run = OxmlElement('w:r')
        text_elem = OxmlElement('w:t')
        text_elem.text = texto
        run.append(text_elem)
        nuevo_p.append(run)
    
    # Insertar el nuevo párrafo después del actual
    parent.insert(parent.index(p_element) + 1, nuevo_p)
    
    return nuevo_p

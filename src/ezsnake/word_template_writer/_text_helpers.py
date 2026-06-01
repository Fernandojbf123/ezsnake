"""
Internal helper functions for text manipulation in Word templates.

This module contains low-level functions for replacing text variables in paragraphs.

Private module - Not intended for direct external use.
Import from the public API in api.py instead.
"""


def replace_text_variables_in_paragraph(paragraph, lista_variables):
    """Reemplaza múltiples marcadores de posición en un párrafo de Word de una sola vez.
    
    Args:
        paragraph: El párrafo donde buscar los marcadores.
        lista_variables: Lista de tuplas (key, value) con los marcadores y sus valores.
                        Ejemplo: [("<<orden_de_servicio>>", "202"), ("<<numero_de_sondas>>", 5)]
    
    Esta función reemplaza todas las variables en un solo paso, evitando problemas
    de estado cuando hay múltiples variables en el mismo párrafo.
    
    Casos especiales:
        - Si el valor es una lista y el marcador es la única variable en el párrafo,
          se crearán múltiples párrafos (uno por cada elemento de la lista).
        - Si hay múltiples variables en el párrafo o hay texto adicional,
          solo se usa el primer elemento de la lista.
    """
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    
    # Obtener el texto completo del párrafo
    full_text = "".join(run.text for run in paragraph.runs)
    
    # Caso especial: Una sola variable con valor tipo lista que ocupa todo el párrafo
    if len(lista_variables) == 1:
        key, value = lista_variables[0]
        # Verificar si es una lista con múltiples elementos y el marcador ocupa todo el párrafo
        if isinstance(value, list) and len(value) > 1 and full_text.strip() == key:
            # Copiar el estilo del párrafo original
            estilo_original = paragraph.style
            formato_original = paragraph.paragraph_format
            
            # Modificar el primer párrafo con el primer elemento
            primer_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            primer_run.text = str(value[0])
            for i in range(1, len(paragraph.runs)):
                paragraph.runs[i].text = ""
            
            # Obtener el elemento actual y su padre
            elemento_actual = paragraph._element
            
            # Insertar párrafos adicionales para el resto de elementos
            for texto in value[1:]:
                # Crear un nuevo elemento de párrafo
                nuevo_elemento = OxmlElement('w:p')
                
                # Insertar el nuevo párrafo después del anterior
                elemento_actual.addnext(nuevo_elemento)
                
                # Crear objeto Paragraph desde el elemento XML
                nuevo_parrafo = Paragraph(nuevo_elemento, paragraph._parent)
                nuevo_parrafo.style = estilo_original
                
                # Agregar el texto al nuevo párrafo
                nuevo_parrafo.add_run(str(texto))
                
                # Actualizar el elemento actual para la siguiente iteración
                elemento_actual = nuevo_elemento
            
            return paragraph
    
    # Comportamiento estándar: reemplazar variables en el texto
    texto_reemplazado = full_text
    for key, value in lista_variables:
        if key in texto_reemplazado:
            # Convertir value a string (si es lista, tomar el primer elemento)
            new_value = str(value[0]) if isinstance(value, list) else str(value)
            # Reemplazar todas las ocurrencias de esta variable
            texto_reemplazado = texto_reemplazado.replace(key, new_value)
    
    # Si no hubo cambios, retornar
    if texto_reemplazado == full_text:
        return paragraph
    
    # Limpiar todos los runs excepto el primero
    primer_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    primer_run.text = texto_reemplazado
    
    # Limpiar el resto de runs
    for i in range(1, len(paragraph.runs)):
        paragraph.runs[i].text = ""
    
    return paragraph


def replace_text_variables_in_tables(doc, diccionario_de_reemplazos):
    """Reemplaza marcadores de posición en todas las celdas de todas las tablas del documento.
    
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Dict con {<<variable>>: valor}
    
    Comportamiento:
        - Itera todas las tablas del documento
        - Busca en todos los párrafos de cada celda
        - Ignora marcadores que inician con: <<fig_, <<ref_, <<tabla_, <<external_doc
        - Reutiliza replace_text_variables_in_paragraph() para preservar formato
    
    Caso de uso típico:
        Tablas semi-estáticas con textos fijos + variables individuales.
        
        Ejemplo en plantilla Word:
        ┌──────────────────────────────┬─────────────────────┐
        │        <<nombre_equipo>>                           │ ← Encabezado
        ├──────────────────────────────┼─────────────────────┤
        │ Serial                       │ <<serial>>          │ ← Fijos + variables
        │ Profundidad máxima           │ <<profundidad>>     │
        └──────────────────────────────┴─────────────────────┘
    
    Nota:
        Complementa a reemplazar_texto_en_plantilla() que solo busca en doc.paragraphs.
        Para tablas dinámicas (llenar con DataFrame completo), usar rellenar_tablas_en_plantilla().
    """
    # Filtrar variables de texto (ignorar fig, ref, tabla, external_doc)
    # Verificar que INICIE con estos prefijos, no solo que los contenga
    prefijos_excluidos = ["<<fig_", "<<ref_", "<<tabla_", "<<external_doc"]
    variables_texto = {
        k: v for k, v in diccionario_de_reemplazos.items() 
        if not any(k.startswith(prefix) for prefix in prefijos_excluidos)
    }
    
    # Iterar todas las tablas del documento
    for table in doc.tables:
        # Iterar todas las celdas de la tabla
        for row in table.rows:
            for cell in row.cells:
                # Cada celda puede tener múltiples párrafos
                for parrafo in cell.paragraphs:
                    # Encontrar variables en este párrafo
                    variables_en_parrafo = []
                    for variable, dato in variables_texto.items():
                        if variable in parrafo.text:
                            variables_en_parrafo.append((variable, dato))
                    
                    # Si hay variables, reemplazarlas (reutiliza función existente)
                    if variables_en_parrafo:
                        replace_text_variables_in_paragraph(parrafo, variables_en_parrafo)
    
    msg = "Se reemplazaron las variables en las tablas del documento."
    print(msg)

"""
Internal helper functions for text manipulation in Word templates.

This module contains low-level functions for replacing text variables in paragraphs.

Private module - Not intended for direct external use.
Import from the public API in api.py instead.
"""

from docx.oxml.ns import qn


def _parse_item_lista_a_texto_y_estilo(item):
    """Normaliza un elemento de lista a (texto, estilo).

    Formatos soportados:
        - "texto"
        - ("texto", "NombreEstilo")
        - ("texto", "") -> usa estilo Normal
    """
    if isinstance(item, tuple):
        texto = str(item[0]) if len(item) > 0 else ""
        estilo = str(item[1]).strip() if len(item) > 1 and item[1] is not None else ""
        return texto, estilo

    return str(item), ""


def _run_contiene_campo_complejo(run):
    """True si el run forma parte de un campo complejo de Word (ej: referencia cruzada).

    Estos runs contienen <w:fldChar> (begin/separate/end) o <w:instrText> y no deben
    ser reescritos: run.text = "" invoca clear_content(), que borra ese contenido y
    rompe el campo (la referencia deja de actualizarse o desaparece).
    """
    r = run._r
    return r.find(qn("w:fldChar")) is not None or r.find(qn("w:instrText")) is not None


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
                    Cada elemento puede ser:
                        * "texto" (usa estilo Normal)
                        * ("texto", "NombreEstilo")
                        * ("", "") para insertar salto de línea con estilo Normal
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
        # Verificar si es una lista y el marcador ocupa todo el párrafo
        if isinstance(value, list) and len(value) > 0 and full_text.strip() == key:
            # Copiar el estilo del párrafo original
            estilo_original = paragraph.style
            estilo_normal = "Normal"

            # Parsear el primer elemento de la lista (texto + estilo opcional)
            texto_inicial, estilo_inicial = _parse_item_lista_a_texto_y_estilo(value[0])
            estilo_a_aplicar = estilo_inicial if estilo_inicial else estilo_normal

            # Aplicar estilo del primer elemento; fallback a estilo original si no existe
            try:
                paragraph.style = estilo_a_aplicar
            except KeyError:
                paragraph.style = estilo_original
            
            # Modificar el primer párrafo con el primer elemento
            primer_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            primer_run.text = texto_inicial
            for i in range(1, len(paragraph.runs)):
                paragraph.runs[i].text = ""
            
            # Obtener el elemento actual y su padre
            elemento_actual = paragraph._element
            
            # Insertar párrafos adicionales para el resto de elementos
            for item in value[1:]:
                texto, estilo = _parse_item_lista_a_texto_y_estilo(item)
                # Crear un nuevo elemento de párrafo
                nuevo_elemento = OxmlElement('w:p')
                
                # Insertar el nuevo párrafo después del anterior
                elemento_actual.addnext(nuevo_elemento)
                
                # Crear objeto Paragraph desde el elemento XML
                nuevo_parrafo = Paragraph(nuevo_elemento, paragraph._parent)
                estilo_a_aplicar = estilo if estilo else estilo_normal
                try:
                    nuevo_parrafo.style = estilo_a_aplicar
                except KeyError:
                    nuevo_parrafo.style = estilo_original
                
                # Agregar el texto al nuevo párrafo
                nuevo_parrafo.add_run(texto)
                
                # Actualizar el elemento actual para la siguiente iteración
                elemento_actual = nuevo_elemento
            
            return paragraph
    
    # Comportamiento estándar: reemplazar variables preservando formato de runs
    
    # Crear diccionario de reemplazos
    reemplazos = {}
    for key, value in lista_variables:
        if isinstance(value, list) and len(value) > 0:
            new_value, _ = _parse_item_lista_a_texto_y_estilo(value[0])
        else:
            new_value = str(value)
        reemplazos[key] = new_value
    
    # Estrategia 1: Intentar reemplazar run por run (caso óptimo - preserva formato)
    for run in paragraph.runs:
        # Los runs que forman parte de un campo complejo (ej: referencias cruzadas
        # <<w:fldChar>>/<<w:instrText>>) no deben tocarse: asignar run.text los destruye,
        # aunque no contengan ningún marcador (run.text = "" también borra fldChar/instrText).
        if _run_contiene_campo_complejo(run):
            continue
        texto_run = run.text
        texto_reemplazado = texto_run
        for key, new_value in reemplazos.items():
            if key in texto_reemplazado:
                texto_reemplazado = texto_reemplazado.replace(key, new_value)
        # Solo reasignar si hubo un cambio real: asignar run.text siempre invoca
        # clear_content(), que borra cualquier contenido no textual del run.
        if texto_reemplazado != texto_run:
            run.text = texto_reemplazado
    
    # Verificar si aún quedan marcadores sin reemplazar (estaban partidos entre runs)
    texto_actual = "".join(run.text for run in paragraph.runs)
    marcadores_pendientes = [key for key in reemplazos.keys() if key in texto_actual]
    
    if not marcadores_pendientes:
        # Todos los reemplazos se hicieron en la Estrategia 1
        return paragraph
    
    # Estrategia 2: Reconstruir preservando formato (para marcadores partidos entre runs)
    # Los runs de campos complejos (referencias cruzadas) actúan como límites fijos:
    # se reconstruye cada tramo de runs "normales" entre ellos de forma aislada e in-place,
    # para no alterar la posición del campo en el párrafo.
    runs = list(paragraph.runs)
    n = len(runs)
    i = 0
    while i < n:
        if _run_contiene_campo_complejo(runs[i]):
            i += 1
            continue
        j = i
        tramo = []
        while j < n and not _run_contiene_campo_complejo(runs[j]):
            tramo.append(runs[j])
            j += 1
        anchor_run = runs[j] if j < n else None
        _reconstruir_tramo_de_runs(paragraph, tramo, reemplazos, anchor_run)
        i = j
    
    return paragraph


def _reconstruir_tramo_de_runs(paragraph, tramo, reemplazos, anchor_run):
    """Reemplaza marcadores partidos entre los runs de `tramo`, preservando su posición.

    `tramo` es una secuencia contigua de runs "normales" (sin fldChar/instrText).
    Los nuevos runs se insertan justo antes de `anchor_run` (o al final del párrafo
    si `anchor_run` es None), y los runs originales del tramo se eliminan.
    """
    texto_original = "".join(run.text for run in tramo)

    texto_nuevo = texto_original
    for key, new_value in reemplazos.items():
        texto_nuevo = texto_nuevo.replace(key, new_value)

    if texto_nuevo == texto_original:
        return

    # Mapear cada posición del texto original a su formato de run
    formato_por_posicion = []
    for run in tramo:
        formato_run = {
            'bold': run.bold,
            'italic': run.italic,
            'underline': run.underline,
            'font_name': run.font.name if run.font.name else None,
            'font_size': run.font.size,
            'color': run.font.color.rgb if run.font.color.rgb else None,
        }
        formato_por_posicion.extend([formato_run] * len(run.text))

    # Construir los nuevos runs con el formato heredado, en el orden correcto
    pos_original = 0
    pos_nueva = 0
    nuevos_runs = []

    while pos_nueva < len(texto_nuevo):
        if pos_original < len(formato_por_posicion):
            formato_actual = formato_por_posicion[pos_original]
        else:
            formato_actual = formato_por_posicion[-1] if formato_por_posicion else {}

        longitud_segmento = 1
        while (pos_nueva + longitud_segmento < len(texto_nuevo) and
               pos_original + longitud_segmento < len(formato_por_posicion) and
               formato_por_posicion[pos_original + longitud_segmento] == formato_actual):
            longitud_segmento += 1

        texto_segmento = texto_nuevo[pos_nueva:pos_nueva + longitud_segmento]

        # add_run() agrega el run al final del párrafo; se reposiciona más abajo.
        new_run = paragraph.add_run(texto_segmento)
        new_run.bold = formato_actual.get('bold')
        new_run.italic = formato_actual.get('italic')
        new_run.underline = formato_actual.get('underline')
        if formato_actual.get('font_name'):
            new_run.font.name = formato_actual.get('font_name')
        if formato_actual.get('font_size'):
            new_run.font.size = formato_actual.get('font_size')
        if formato_actual.get('color'):
            new_run.font.color.rgb = formato_actual.get('color')
        nuevos_runs.append(new_run)

        pos_nueva += longitud_segmento
        pos_original += longitud_segmento

    # Reubicar los nuevos runs justo antes del límite (o dejarlos al final si no hay límite)
    if anchor_run is not None:
        anchor_element = anchor_run._r
        for new_run in nuevos_runs:
            anchor_element.addprevious(new_run._r)

    # Eliminar los runs originales del tramo (ya reemplazados por los nuevos)
    for run in tramo:
        run._r.getparent().remove(run._r)


def replace_text_variables_in_tables(doc, diccionario_de_reemplazos):
    """Reemplaza marcadores de posición en todas las celdas de todas las tablas del documento.
    
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Dict con {<<variable>>: valor}
    
    Comportamiento:
        - Itera todas las tablas del documento
        - Busca en todos los párrafos de cada celda
        - Ignora marcadores que inician con: <<fig, <<reffigura, <<reftabla_, <<nuevatabla_, <<editartabla, <<external_doc
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
    # Filtrar variables de texto (ignorar fig, reffigura, tabla, external_doc)
    # Verificar que INICIE con estos prefijos, no solo que los contenga
    prefijos_excluidos = [
        "<<fig",
        "<<reffigura",
        "<<reftabla_",
        "<<nuevatabla_",
        "<<editartabla",
        "<<external_doc",
    ]
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

"""
Public API for Word Template Writer Module
==========================================

This module provides high-level orchestrator functions (in Spanish) for manipulating
Word templates. These are the main functions users should call.

Functions:
    - reemplazar_variable_por_figura: Insert figures with/without captions
    - reemplazar_referencias_cruzadas_de_figuras: Create cross-references to figures
    - reemplazar_variable_por_tabla: Fill table placeholders and prepare table references
    - reemplazar_referencias_cruzadas_de_tablas: Create cross-references to tables
    - reemplazar_texto_en_plantilla: Replace text variables in template
    - insertar_lista_en_plantilla: Insert bulleted lists in template
    - insertar_documento_externo_en_plantilla: Insert external Word documents
    - rellenar_tablas_en_plantilla: Fill tables with DataFrame data
"""

# Import helper functions from private modules
from ._figure_helpers import (
    aux_insertar_figura_sin_titulo,
    aux_insertar_figuras_con_titulo,
    construir_elementos_referencia_cruzada,
)
from ._text_helpers import replace_text_variables_in_paragraph, _run_contiene_campo_complejo
from ._document_helpers import insert_external_document
from ._table_helpers import (
    fill_table,
    procesar_reemplazar_variable_por_tabla,
    procesar_rellenar_tablas_en_plantilla,
    reemplazar_variables_en_tablas_del_documento,
)


def reemplazar_variable_por_figura(doc, diccionario_de_reemplazos: dict):
    """Inserta todas las figuras definidas en el diccionario en la plantilla de Word.
    
    Muta el diccionario de entrada al unirlo con la información de los bookmarks 
    creados para referencias cruzadas.
        
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario donde las claves que comienzan con "<<fig" 
                                   contienen listas de figuras a insertar.
    
    Returns:
        None (el diccionario de entrada se muta in-place agregando keys "<<reffigura*>>")
    
    Casos soportados:
        - Caso 1: Figuras SIN título - titulo="" y bookmark=""
                 Se insertan imágenes sin Caption, numeración ni bookmarks
        - Caso 2: Figuras CON título - titulo y bookmark con contenido
                 Se crean pies de figura con estilo Caption, campo SEQ y bookmarks
    
    Estructura esperada del diccionario:
        {
            # Caso 1: Figuras SIN título
            "<<fig_fotos>>": [
                {"ruta": "foto1.png", "titulo": "", "tamanio": 6, "bookmark": "", "estilo_figura": "Figura", "estilo_titulo": "Normal"},
                {"ruta": "foto2.png", "titulo": "", "tamanio": 5, "bookmark": "", "estilo_figura": "Figura", "estilo_titulo": "Normal"}
            ],
            
            # Caso 2: Figuras CON título
            "<<fig_mapas>>": [
                {"ruta": "mapa1.png", "titulo": "Mapa de ubicación", "tamanio": 6, "bookmark": "RefFigura_Mapa1", "estilo_figura": "Figura", "estilo_titulo": "Normal"},
                {"ruta": "mapa2.png", "titulo": "Temperatura del agua", "tamanio": 5, "bookmark": "RefFigura_Temp", "estilo_figura": "Figura", "estilo_titulo": "Normal"}
            ],
            
            "<<orden_servicio>>": "12345",  # Variables no-figura se ignoran aquí
        }
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import reemplazar_variable_por_figura, reemplazar_referencias_cruzadas_de_figuras
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<fig_mapas>>": [
                {"ruta": "mapa1.png", "titulo": "Ubicación sondas", "tamanio": 6, "bookmark": "RefFigura_Mapa1", "estilo_figura": "Figura", "estilo_titulo": "Normal"},
                {"ruta": "mapa2.png", "titulo": "Temperatura", "tamanio": 5, "bookmark": "RefFigura_Temp", "estilo_figura": "Figura", "estilo_titulo": "Normal"}
            ],
            "<<fig_fotos>>": [
                {"ruta": "foto1.png", "titulo": "", "tamanio": 4, "bookmark": "", "estilo_figura": "Figura", "estilo_titulo": "Normal"}
            ]
        }
        
        # Paso 1: Insertar las figuras (agrega keys <<reffigura*>> al diccionario)
        reemplazar_variable_por_figura(doc, diccionario)
        # El diccionario ahora contiene:
        # {
        #     "<<reffigura_mapas>>": ["RefFigura_Mapa1", "RefFigura_Temp"],
        #     "<<reffigura_fotos>>": None,
        #     ... (keys originales se mantienen)
        # }
        
        # Paso 2: Insertar referencias cruzadas (si hay marcadores <<reffigura*>> en la plantilla)
        # Ejemplo: "De la <<reffigura_mapas>> se observa..." → "De la Figura 1 a la 2 se observa..."
        reemplazar_referencias_cruzadas_de_figuras(doc, diccionario)
        
        doc.save('documento_con_figuras.docx')
    """
    bookmarks_info = {}
    
    if diccionario_de_reemplazos is None:
        raise ValueError("El diccionario de reemplazos no puede ser None.")
    
    # Filtrar solo las variables que son figuras (comienzan con "<<fig")
    variables_figuras = {k: v for k, v in diccionario_de_reemplazos.items() if k.startswith("<<fig")}
    
    # Procesar cada variable de figura
    for variable, datos_figuras in variables_figuras.items():
        if not isinstance(datos_figuras, list):
            print(f"Advertencia: La variable '{variable}' no contiene una lista. Se omite.")
            continue
        
        if len(datos_figuras) == 0:
            print(f"Advertencia: La variable '{variable}' contiene una lista vacía. Se omite.")
            continue
        
        # Determinar si son figuras CON o SIN título
        # Caso 1: Figuras SIN título - todos los títulos están vacíos
        tiene_titulo = any(item.get("titulo", "") != "" for item in datos_figuras)
        
        # Buscar el marcador en todos los párrafos del documento
        for parrafo in doc.paragraphs:
            if variable in parrafo.text:
                
                # Crear key de referencia: "<<fig_mapas>>" -> "<<reffigura_mapas>>"
                variable_ref_key = variable.replace("<<fig", "<<reffigura", 1)
                
                if not tiene_titulo:
                    # Caso 1: Figuras SIN título (no Caption, no bookmark)
                    resultado = aux_insertar_figura_sin_titulo(parrafo, variable, datos_figuras)
                    if resultado:
                        bookmarks_info[variable_ref_key] = None  # Sin bookmarks para figuras sin título
                        break  # Ya se insertó, pasar a la siguiente variable
                else:
                    # Caso 2: Figuras CON título (Caption + SEQ + Bookmarks)
                    bookmarks_creados = aux_insertar_figuras_con_titulo(parrafo, variable, datos_figuras)
                    if bookmarks_creados:
                        bookmarks_info[variable_ref_key] = bookmarks_creados
                        break  # Ya se insertó, pasar a la siguiente variable
    
    # Mutar el diccionario agregando información de bookmarks
    diccionario_de_reemplazos.update(bookmarks_info)
    
    msg = "Figuras insertadas y bookmarks creados para referencias cruzadas"
    print(msg)


def reemplazar_referencias_cruzadas_de_figuras(doc, diccionario_de_reemplazos: dict):
    """Reemplaza marcadores <<reffigura*>> con referencias cruzadas a figuras.
    
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario retornado/mutado por reemplazar_variable_por_figura
                      Formato: {"<<reffigura_demo>>": ["RefFigura_Mapa1", "RefFigura_Temp"], ...}
    
    Comportamiento:
        - Busca variables que empiecen con "<<reffigura" en los párrafos
        - Si la lista tiene 1 bookmark: inserta "Figura X"
        - Si la lista tiene 2+ bookmarks: inserta "Figura X a la Y"
        - X e Y son referencias cruzadas reales (campos REF) que muestran solo el número
        - Procesa TODOS los marcadores de un párrafo en una sola pasada
    
    Nota importante:
        Los bookmarks creados por crear_pie_de_figura solo incluyen el número de la figura,
        no el texto "Figura" ni el título. Por eso las referencias muestran solo el número.
    
    Ejemplo:
        Entrada en plantilla: "De la <<reffigura_demo_con_titulo>> se muestra el poder."
        Salida: "De la Figura 8 a la 10 se muestra el poder."
                (donde "8" y "10" son campos REF clickeables que muestran solo el número)
    """
    # Filtrar solo variables con bookmarks válidos
    variables_validas = {k: v for k, v in diccionario_de_reemplazos.items() 
                        if v is not None and k.startswith("<<reffigura")}
    
    def _piezas(lista_bookmarks):
        primer_bookmark = lista_bookmarks[0]
        ultimo_bookmark = lista_bookmarks[-1]
        if len(lista_bookmarks) == 1:
            # Caso: Solo una figura - "Figura X"
            return [("ref", primer_bookmark, "Figura")]
        # Caso: Múltiples figuras - "Figura X a la Y"
        return [
            ("ref", primer_bookmark, "Figura"),
            ("text", " a la "),
            ("ref", ultimo_bookmark, "Figura"),
        ]

    for parrafo in doc.paragraphs:
        _procesar_marcadores_referencia_en_parrafo(parrafo, variables_validas, _piezas)

    msg = "Referencias cruzadas insertadas."
    print(msg)


def _procesar_marcadores_referencia_en_parrafo(parrafo, variables_validas, obtener_piezas):
    """Reemplaza marcadores <<ref*>> por referencias cruzadas, sin tocar el resto del párrafo.

    A diferencia de una reconstrucción total del párrafo, esto preserva:
        - El formato (bold/italic/etc.) de los runs que no contienen marcadores.
        - Cualquier otro campo complejo ya presente (ej: otra referencia cruzada),
          tratándolo como un límite fijo que no se debe borrar ni mover.
    """
    runs = list(parrafo.runs)
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
        _reconstruir_tramo_con_referencias(parrafo, tramo, variables_validas, obtener_piezas, anchor_run)
        i = j


def _reconstruir_tramo_con_referencias(paragraph, tramo, variables_validas, obtener_piezas, anchor_run):
    """Reemplaza marcadores dentro de `tramo` (runs "normales" contiguos) por referencias cruzadas.

    Preserva el formato original de cada segmento de texto y respeta la posición del
    tramo en el párrafo (los nuevos elementos se insertan justo antes de `anchor_run`).
    """
    texto_original = "".join(run.text for run in tramo)

    # Encontrar todas las ocurrencias de marcadores en este tramo, ordenadas por posición
    ocurrencias = []
    for variable_ref, lista_bookmarks in variables_validas.items():
        start = 0
        while True:
            pos = texto_original.find(variable_ref, start)
            if pos == -1:
                break
            ocurrencias.append((pos, variable_ref, lista_bookmarks))
            start = pos + len(variable_ref)

    if not ocurrencias:
        return  # Nada que reemplazar: se deja el tramo intacto

    ocurrencias.sort(key=lambda x: x[0])

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

    def _formato_en(pos):
        if not formato_por_posicion:
            return {}
        idx = min(max(pos, 0), len(formato_por_posicion) - 1)
        return formato_por_posicion[idx]

    nuevos_elementos = []

    def _agregar_texto(texto_segmento, formato):
        if not texto_segmento:
            return
        # add_run() agrega el run al final del párrafo; se reposiciona más abajo.
        new_run = paragraph.add_run(texto_segmento)
        new_run.bold = formato.get('bold')
        new_run.italic = formato.get('italic')
        new_run.underline = formato.get('underline')
        if formato.get('font_name'):
            new_run.font.name = formato.get('font_name')
        if formato.get('font_size'):
            new_run.font.size = formato.get('font_size')
        if formato.get('color'):
            new_run.font.color.rgb = formato.get('color')
        nuevos_elementos.append(new_run._r)

    pos_actual = 0

    for pos_marcador, variable_ref, lista_bookmarks in ocurrencias:
        if pos_marcador > pos_actual:
            _agregar_texto(texto_original[pos_actual:pos_marcador], _formato_en(pos_actual))

        for pieza in obtener_piezas(lista_bookmarks):
            if pieza[0] == "text":
                _agregar_texto(pieza[1], _formato_en(pos_marcador))
            else:
                _, bookmark, texto_antes_pieza = pieza
                nuevos_elementos.extend(
                    construir_elementos_referencia_cruzada(bookmark, texto_antes=texto_antes_pieza, mostrar_numero=True)
                )

        pos_actual = pos_marcador + len(variable_ref)

    if pos_actual < len(texto_original):
        _agregar_texto(texto_original[pos_actual:], _formato_en(pos_actual))

    # Reubicar los nuevos elementos justo antes del límite (o dejarlos al final si no hay límite)
    if anchor_run is not None:
        anchor_element = anchor_run._r
        for elem in nuevos_elementos:
            anchor_element.addprevious(elem)
    else:
        for elem in nuevos_elementos:
            paragraph._p.append(elem)

    # Eliminar los runs originales del tramo (ya reemplazados por los nuevos)
    for run in tramo:
        run._r.getparent().remove(run._r)


def reemplazar_referencias_cruzadas_de_tablas(doc, diccionario_de_reemplazos: dict):
    """Reemplaza marcadores <<refnuevatabla_*>> con referencias cruzadas a tablas.

    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario mutado por reemplazar_variable_por_tabla
                                  Formato: {"<<refnuevatabla_demo>>": ["RefTabla_MiTabla"], ...}

    Comportamiento:
        - Busca variables que empiecen con "<<refnuevatabla_" en los párrafos
        - Si la lista tiene 1 bookmark: inserta "Tabla X"
        - Si la lista tiene 2+ bookmarks: inserta "Tabla X a la Y"
        - X e Y son referencias cruzadas reales (campos REF) que muestran el número/etiqueta
        - Procesa TODOS los marcadores de un párrafo en una sola pasada

    Ejemplo de uso:
        from docx import Document
        from word_template_writer import reemplazar_referencias_cruzadas_de_tablas

        doc = Document('plantilla.docx')
        diccionario = {
            "<<refnuevatabla_resultados>>": ["RefTabla_Resultados_1"],
            "<<refnuevatabla_series>>": ["RefTabla_Serie_A", "RefTabla_Serie_B"],
        }

        reemplazar_referencias_cruzadas_de_tablas(doc, diccionario)
        doc.save('documento_con_refs_tablas.docx')
    """
    variables_validas = {
        k: v for k, v in diccionario_de_reemplazos.items()
        if v is not None and (k.startswith("<<refnuevatabla_") or k.startswith("<<reftabla_"))
    }

    def _piezas(lista_bookmarks):
        primer_bookmark = lista_bookmarks[0]
        ultimo_bookmark = lista_bookmarks[-1]
        if len(lista_bookmarks) == 1:
            return [("ref", primer_bookmark, "Tabla")]
        return [
            ("ref", primer_bookmark, "Tabla"),
            ("text", " a la "),
            ("ref", ultimo_bookmark, ""),
        ]

    for parrafo in doc.paragraphs:
        _procesar_marcadores_referencia_en_parrafo(parrafo, variables_validas, _piezas)

    msg = "Referencias cruzadas de tablas insertadas."
    print(msg)


def reemplazar_variable_por_tabla(doc, diccionario_de_reemplazos: dict):
    """Reemplaza variables <<nuevatabla_*>> por tablas y prepara referencias <<refnuevatabla_*>>.

    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario con variables del documento

    Estructura esperada de cada variable de tabla:
        {
            "tabla": pd.DataFrame,
            "estilos_de_tabla": dict,
            "bookmark": "RefTabla_MiTabla"
            "titulo": "Texto del título de tabla",
            "estilo_de_titulo": "Carcentrado"
        }

    Comportamiento:
        - Busca keys que comienzan con "<<nuevatabla_"
        - Crea la tabla desde cero donde aparece el marcador de párrafo
        - Inserta un título numerado de Word con campo SEQ Tabla y bookmark
        - El título usa estilo Caption para aparecer en el índice de tablas
        - Muta diccionario_de_reemplazos agregando "<<refnuevatabla_*>>": [bookmark]

    Ejemplo de uso:
        import pandas as pd
        from docx import Document
        from word_template_writer import EstilosTabla, reemplazar_variable_por_tabla

        doc = Document('plantilla.docx')
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        estilos = EstilosTabla(doc).to_dict()

        diccionario = {
            "<<nuevatabla_resultados>>": {
                "tabla": df,
                "estilos_de_tabla": estilos,
                "bookmark": "RefTabla_Resultados_1",
                "titulo": "Resultados del análisis",
                "estilo_de_titulo": "Carcentrado"
            }
        }

        reemplazar_variable_por_tabla(doc, diccionario)
        doc.save('documento_con_tablas.docx')
    """
    procesar_reemplazar_variable_por_tabla(doc, diccionario_de_reemplazos)

    msg = "Tablas insertadas y referencias de tabla preparadas en el diccionario."
    print(msg)


def reemplazar_texto_en_plantilla(doc, diccionario_de_reemplazos):
    """Reemplaza los marcadores de posición de texto en un documento de Word.
    
    Args:
        doc: El documento de Word (objeto Document de python-docx).
        diccionario_de_reemplazos: Diccionario donde las claves son los marcadores de posición 
                                  a buscar (por ejemplo, "<<orden_de_servicio>>") y los valores 
                                  son los textos que los reemplazarán (ej: "12345").
                                  Los valores pueden ser strings o listas de strings.
    
        Comportamiento:
                - Ignora marcadores reservados para:
                    * Figuras (<<fig...>>)
                    * Referencias de figuras (<<reffigura...>>)
                    * Tablas (<<nuevatabla_...>>, <<editartabla...>>)
                    * Referencias de tablas (<<refnuevatabla_...>>, <<reftabla_...>>)
                    * Documentos externos (<<external_doc_...>>)
        - Procesa todas las variables de texto en cada párrafo de una sola vez
        - Preserva el formato del primer run del párrafo
        - Si el valor es una lista y el marcador ocupa todo el párrafo:
                    * Se crean múltiples párrafos (uno por cada elemento de la lista)
                    * El valor debe ser una lista de tuplas: (texto, estilo)
                    * Si estilo == "", se aplica estilo "Normal"
                    * Si texto == "", se inserta un salto de línea (párrafo vacío)
        - Si hay texto adicional o múltiples marcadores en el mismo párrafo:
          * Solo se usa el primer elemento de la lista
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import reemplazar_texto_en_plantilla
        
        doc = Document('plantilla.docx')
        
        # Caso 1: Reemplazo simple con strings
        diccionario = {
            "<<orden_servicio>>": "12345",
            "<<cliente>>": "ACME Corporation",
            "<<fecha>>": "21/05/2026",
        }
        
        # Caso 2: Reemplazo con múltiples párrafos (lista de tuplas)
        diccionario = {
            "<<seccion_1>>": [
                ("este es el primer párrafo", "estilo 2"),
                ("este es el segundo párrafo", ""),
                ("", ""),
                ("tercer párrafo", "estilo 3"),
            ],
            "<<orden_servicio>>": "12345",
        }
        
        reemplazar_texto_en_plantilla(doc, diccionario)
        doc.save('documento_con_texto.docx')
    
    Nota:
        Para que se creen múltiples párrafos, el marcador debe ser el único
        contenido del párrafo en la plantilla (sin texto adicional).
    """
    # Filtrar solo variables de texto, excluyendo prefijos reservados de otras operaciones.
    prefijos_reservados = (
        "<<fig",
        "<<reffigura",
        "<<nuevatabla_",
        "<<editartabla",
        "<<refnuevatabla_",
        "<<reftabla_",
        "<<external_doc_",
        "<<lista_",
        "<<ul_",
    )
    variables_texto = {
        k: v for k, v in diccionario_de_reemplazos.items()
        if not any(k.startswith(prefijo) for prefijo in prefijos_reservados)
    }
    
    # Para cada párrafo, procesar TODAS las variables de texto de una sola vez
    for it in [0,1]:
        for parrafo in doc.paragraphs:
            # Encontrar todas las variables que están en este párrafo
            variables_en_parrafo = []
            for variable, dato in variables_texto.items():
                if variable in parrafo.text:
                    variables_en_parrafo.append((variable, dato))
            
            # Si hay variables en este párrafo, reemplazarlas todas de una vez
            if variables_en_parrafo:
                replace_text_variables_in_paragraph(parrafo, variables_en_parrafo)
    
    msg = "Se agregaron los textos al documento."
    print(msg)


def insertar_lista_en_plantilla(doc, diccionario_de_reemplazos, nombre_estilo="List Bullet"):
    """Inserta listas con viñetas en la plantilla de Word.
    
    Args:
        doc: El documento de Word (objeto Document de python-docx).
        diccionario_de_reemplazos: Diccionario donde las claves que comienzan con "<<lista_" 
                                   contienen listas de elementos a insertar como viñetas.
        nombre_estilo: Nombre del estilo de párrafo a aplicar a cada elemento de la lista.
                      Por defecto usa "List Bullet" (estilo incorporado de Word).
                      Se recomienda crear un estilo personalizado llamado "unsorted_list" 
                      en la plantilla para mayor control del formato.
    
    Comportamiento:
        - Busca marcadores que empiecen con "<<lista_" en los párrafos
        - El marcador DEBE estar solo en el párrafo (sin texto adicional)
        - Si el marcador no está solo, lanza ValueError
        - Reemplaza el párrafo completo con múltiples párrafos (uno por elemento)
        - Aplica el estilo especificado a cada párrafo nuevo
        - Preserva la posición original del marcador en el documento
    
    Estructura esperada del diccionario:
        {
            "<<lista_resultados>>": [
                "Primer resultado del análisis",
                "Segundo resultado encontrado",
                "Tercer resultado importante"
            ],
            "<<lista_recomendaciones>>": [
                "Primera recomendación",
                "Segunda recomendación"
            ],
            "<<orden_servicio>>": "12345"  # Variables no-lista se ignoran aquí
        }
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import insertar_lista_en_plantilla
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<lista_hallazgos>>": [
                "Se observó un incremento del 15% en la temperatura",
                "La salinidad mostró valores consistentes con el promedio histórico",
                "Se detectaron concentraciones elevadas de clorofila-a"
            ],
            "<<lista_equipos>>": [
                "CTD SBE 911plus",
                "ADCP Workhorse 300 kHz",
                "Fluorómetro WETLabs ECO"
            ]
        }
        
        # Usando estilo por defecto
        insertar_lista_en_plantilla(doc, diccionario)
        
        # Usando estilo personalizado (debe existir en la plantilla)
        insertar_lista_en_plantilla(doc, diccionario, nombre_estilo="unsorted_list")
        
        doc.save('documento_con_listas.docx')
    
    Raises:
        ValueError: Si el marcador no está solo en el párrafo o si el valor no es una lista
    
    Nota importante:
        Para que esta función trabaje correctamente, el marcador <<lista_*>> debe ser
        el ÚNICO contenido del párrafo en la plantilla. Por ejemplo:
        
        ✓ CORRECTO:   Párrafo que solo contiene "<<lista_hallazgos>>"
        ✗ INCORRECTO: "Los hallazgos son: <<lista_hallazgos>>"
    """
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    
    # Filtrar solo variables que son listas (comienzan con "<<lista_")
    variables_listas = {k: v for k, v in diccionario_de_reemplazos.items() if k.startswith("<<lista_")}
    
    # Procesar cada variable de lista
    for variable, elementos_lista in variables_listas.items():
        # Validar que el valor sea una lista
        if not isinstance(elementos_lista, list):
            raise ValueError(
                f"El valor de '{variable}' debe ser una lista. "
                f"Se recibió: {type(elementos_lista).__name__}"
            )
        
        if len(elementos_lista) == 0:
            print(f"Advertencia: La variable '{variable}' contiene una lista vacía. Se omite.")
            continue
        
        # Buscar el marcador en todos los párrafos del documento
        for parrafo in doc.paragraphs:
            if variable in parrafo.text:
                # Obtener el texto completo del párrafo
                full_text = "".join(run.text for run in parrafo.runs).strip()
                
                # Validar que el marcador esté SOLO en el párrafo
                if full_text != variable:
                    raise ValueError(
                        f"El marcador '{variable}' debe estar solo en el párrafo. "
                        f"Texto encontrado: '{full_text}'. "
                        f"Los elementos <<lista_*>> deben estar solos en un párrafo."
                    )
                
                # Reemplazar el primer párrafo con el primer elemento de la lista
                primer_run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
                primer_run.text = str(elementos_lista[0])
                
                # Limpiar el resto de runs del párrafo
                for i in range(1, len(parrafo.runs)):
                    parrafo.runs[i].text = ""
                
                # Aplicar el estilo al primer párrafo
                try:
                    parrafo.style = nombre_estilo
                except KeyError:
                    print(f"Advertencia: El estilo '{nombre_estilo}' no existe en el documento. "
                          f"Se mantiene el estilo original del párrafo.")
                
                # Obtener el elemento actual y su padre
                elemento_actual = parrafo._element
                
                # Insertar párrafos adicionales para el resto de elementos
                for texto in elementos_lista[1:]:
                    # Crear un nuevo elemento de párrafo
                    nuevo_elemento = OxmlElement('w:p')
                    
                    # Insertar el nuevo párrafo después del anterior
                    elemento_actual.addnext(nuevo_elemento)
                    
                    # Crear objeto Paragraph desde el elemento XML
                    nuevo_parrafo = Paragraph(nuevo_elemento, parrafo._parent)
                    
                    # Aplicar el estilo
                    try:
                        nuevo_parrafo.style = nombre_estilo
                    except KeyError:
                        pass  # Si falla, usa el estilo por defecto
                    
                    # Agregar el texto al nuevo párrafo
                    nuevo_parrafo.add_run(str(texto))
                    
                    # Actualizar el elemento actual para la siguiente iteración
                    elemento_actual = nuevo_elemento
                
                # Ya se procesó este marcador, pasar al siguiente
                break
    
    msg = "Listas insertadas en el documento."
    print(msg)


def insertar_documento_externo_en_plantilla(doc, diccionario_de_reemplazos):
    """Inserta uno o más documentos Word externos en la plantilla.
    
    Args:
        doc: El documento de Word (objeto Document de python-docx).
        diccionario_de_reemplazos: Diccionario que debe contener keys que inicien con "<<external_doc_"
                                  con el valor siendo una ruta (string) o lista de rutas a 
                                  documentos Word externos.
    
    Comportamiento:
        - Busca marcadores "<<external_doc_*>>" en los párrafos
        - Inserta el contenido completo de cada documento externo (párrafos y tablas)
        - Copia imágenes y mantiene relaciones correctas
        - Excluye headers, footers, marcas de agua y configuraciones de sección
        - Agrega saltos de página entre múltiples documentos
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import insertar_documento_externo_en_plantilla
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<external_doc_plan_de_crucero>>": "plan_crucero.docx",
            # O múltiples documentos:
            # "<<external_doc_plan_de_crucero>>": ["plan1.docx", "plan2.docx"],
        }
        
        insertar_documento_externo_en_plantilla(doc, diccionario)
        doc.save('documento_con_plan.docx')
    """
     # Filtrar solo variables que NO son figuras, referencias o documentos externos
    variables_texto = {k: v for k, v in diccionario_de_reemplazos.items() if k.startswith("<<external_doc_")}
    
    # Para cada párrafo, procesar TODAS las variables de texto de una sola vez
    for parrafo in doc.paragraphs:
        # Encontrar todas las variables que están en este párrafo
        variables_en_parrafo = []
        for variable, dato in variables_texto.items():
            if variable in parrafo.text:
                variables_en_parrafo.append((variable, dato))        
                insert_external_document(parrafo, variable, dato, doc)
            
                msg = f"Documento insertado {variable}"
                if isinstance(dato, list) and len(dato) > 1:
                    msg = "Planes de crucero insertados"
                print(msg)


def rellenar_tablas_en_plantilla(doc, diccionario_de_reemplazos: dict):      
    """Rellena tablas en la plantilla usando keys <<editartabla...>>.
    
    Soporta estilos configurables, colores de fondo, MultiIndex, y merge de celdas.
    
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario con entradas que inician con <<editartabla
            Estructura por key:
            {
                "tabla": pd.DataFrame,
                "estilos_de_tabla": EstilosTabla | dict | None,
                "opciones_de_tabla": OpcionesTabla | dict | None,
            }
    
    Comportamiento:
        - Busca el marcador en cualquier celda de las tablas del documento
        - Inserta datos del DataFrame fila por fila, respetando orden de columnas
        - Aplica estilos según jerarquía: celda > fila > columna > defecto
        - Soporta colores de fondo RGB y estilos de párrafo del documento
        - Aplana MultiIndex automáticamente si está activado en opciones
        - Detecta y combina celdas verticales con valores repetidos (si está activado)
        - Elimina fila marcador automáticamente (si está activado)
    
    Ejemplo de uso:
        >>> import pandas as pd
        >>> from docx import Document
        >>> from word_template_writer import EstilosTabla, OpcionesTabla, rellenar_tablas_en_plantilla
        >>>
        >>> doc = Document('plantilla.docx')
        >>> estilos = EstilosTabla(doc)
        >>> opciones = OpcionesTabla()
        >>> df = pd.DataFrame({"Secuencia": [0, 1], "Localización": ["BOT-01", "BOT-02"]})
        >>>
        >>> diccionario = {
        ...     "<<editartabla_posiciones>>": {
        ...         "tabla": df,
        ...         "estilos_de_tabla": estilos,
        ...         "opciones_de_tabla": opciones,
        ...     }
        ... }
        >>>
        >>> rellenar_tablas_en_plantilla(doc, diccionario)
        >>> doc.save('documento_con_tablas_editadas.docx')
    
    Raises:
        ValueError: Si el marcador no se encuentra, si hay errores en la configuración,
                   o si los datos están vacíos
    
    Nota:
        - La plantilla debe tener un marcador <<nombre>> en alguna celda de la tabla
        - El marcador se coloca típicamente en la primera fila de datos (no encabezado)
        - Los datos se insertan por ORDEN de columna, no por nombre
        - Columna 0 del DataFrame → Columna 0 de la tabla
    """
    
    procesar_rellenar_tablas_en_plantilla(doc, diccionario_de_reemplazos)

    msg = "Tablas rellenadas correctamente."
    print(msg)


def reemplazar_variable_en_tabla(doc, diccionario_de_reemplazos):
    """Alias singular de reemplazar_variables_en_tablas para compatibilidad."""
    reemplazar_variables_en_tablas(doc, diccionario_de_reemplazos)


def reemplazar_variables_en_tablas(doc, diccionario_de_reemplazos):
    """Reemplaza marcadores de posición en celdas de tablas del documento.
    
    Complementa a reemplazar_texto_en_plantilla() para tablas semi-estáticas
    donde solo se necesita reemplazar variables individuales, NO llenar
    toda la tabla con un DataFrame.
    
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Dict donde keys son marcadores (ej: "<<serial>>")
                                  y values son los textos de reemplazo
    
    Comportamiento:
        - Busca en TODAS las tablas del documento automáticamente
        - Ignora variables que inician con: <<fig, <<reffigura, <<refnuevatabla_, <<reftabla_, <<nuevatabla_, <<editartabla, <<external_doc
        - Preserva el formato del texto original
        - Reutiliza la misma lógica de replace_text_variables_in_paragraph()
    
    Caso de uso típico:
        Tabla con:
        - Textos fijos en columna 0: "Serial", "Profundidad máxima", "Precisión"
        - Variables en columna 1: <<serial>>, <<profundidad>>, <<precision>>
        - Encabezado con variable: <<nombre_equipo>>
    
    Ejemplo:
        from docx import Document
        from word_template_writer import reemplazar_variables_en_tablas
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<nombre_equipo>>": "Sonda CTD #1",
            "<<serial>>": "4878505",
            "<<profundidad>>": "200 m",
            "<<precision>>": "±0.01°C"
        }
        
        reemplazar_variables_en_tablas(doc, diccionario)
        doc.save('resultado.docx')
    
    Ver también:
        - reemplazar_texto_en_plantilla(): para párrafos del documento (NO tablas)
        - rellenar_tablas_en_plantilla(): para tablas dinámicas con DataFrames
    """
    reemplazar_variables_en_tablas_del_documento(doc, diccionario_de_reemplazos)

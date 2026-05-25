"""
Public API for Word Template Writer Module
==========================================

This module provides high-level orchestrator functions (in Spanish) for manipulating
Word templates. These are the main functions users should call.

Functions:
    - insertar_figuras_en_plantilla: Insert figures with/without captions
    - insertar_referencias_cruzadas_en_plantilla: Create cross-references to figures
    - reemplazar_texto_en_plantilla: Replace text variables in template
    - insertar_documento_externo_en_plantilla: Insert external Word documents
    - rellenar_tablas_en_plantilla: Fill tables with DataFrame data
"""

# Import helper functions from private modules
from ._figure_helpers import (
    aux_insertar_figura_sin_titulo,
    aux_insertar_figuras_con_titulo,
    aux_insertar_referencia_cruzada,
)
from ._text_helpers import replace_text_variables_in_paragraph, replace_text_variables_in_tables
from ._document_helpers import insert_external_document
from ._table_helpers import fill_table


def insertar_figuras_en_plantilla(doc, diccionario_de_reemplazos: dict):
    """Inserta todas las figuras definidas en el diccionario en la plantilla de Word.
    
    Muta el diccionario de entrada al unirlo con la información de los bookmarks 
    creados para referencias cruzadas.
        
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario donde las claves que comienzan con "<<fig_" 
                                   contienen listas de figuras a insertar.
    
    Returns:
        None (el diccionario se muta in-place agregando keys "<<ref_*>>")
    
    Casos soportados:
        - Caso 1: Figuras SIN título - titulo="" y bookmark=""
                 Se insertan imágenes sin Caption, numeración ni bookmarks
        - Caso 2: Figuras CON título - titulo y bookmark con contenido
                 Se crean pies de figura con estilo Caption, campo SEQ y bookmarks
    
    Estructura esperada del diccionario:
        {
            # Caso 1: Figuras SIN título
            "<<fig_fotos>>": [
                {"ruta": "foto1.png", "titulo": "", "tamanio": 6, "bookmark": ""},
                {"ruta": "foto2.png", "titulo": "", "tamanio": 5, "bookmark": ""}
            ],
            
            # Caso 2: Figuras CON título
            "<<fig_mapas>>": [
                {"ruta": "mapa1.png", "titulo": "Mapa de ubicación", "tamanio": 6, "bookmark": "_Ref_Mapa1"},
                {"ruta": "mapa2.png", "titulo": "Temperatura del agua", "tamanio": 5, "bookmark": "_Ref_Temp"}
            ],
            
            "<<orden_servicio>>": "12345",  # Variables no-figura se ignoran aquí
        }
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import insertar_figuras_en_plantilla, insertar_referencias_cruzadas_en_plantilla
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<fig_mapas>>": [
                {"ruta": "mapa1.png", "titulo": "Ubicación sondas", "tamanio": 6, "bookmark": "_Ref_Mapa1"},
                {"ruta": "mapa2.png", "titulo": "Temperatura", "tamanio": 5, "bookmark": "_Ref_Temp"}
            ],
            "<<fig_fotos>>": [
                {"ruta": "foto1.png", "titulo": "", "tamanio": 4, "bookmark": ""}
            ]
        }
        
        # Paso 1: Insertar las figuras (agrega keys <<ref_*>> al diccionario)
        insertar_figuras_en_plantilla(doc, diccionario)
        # El diccionario ahora contiene:
        # {
        #     "<<ref_mapas>>": ["_Ref_Mapa1", "_Ref_Temp"],
        #     "<<ref_fotos>>": None,
        #     ... (keys originales se mantienen)
        # }
        
        # Paso 2: Insertar referencias cruzadas (si hay marcadores <<ref_*>> en la plantilla)
        # Ejemplo: "De la <<ref_mapas>> se observa..." → "De la Figura 1 a la 2 se observa..."
        insertar_referencias_cruzadas_en_plantilla(doc, diccionario)
        
        doc.save('documento_con_figuras.docx')
    """
    bookmarks_info = {}
    
    if diccionario_de_reemplazos is None:
        raise ValueError("El diccionario de reemplazos no puede ser None.")
    
    # Filtrar solo las variables que son figuras (comienzan con "<<fig_")
    variables_figuras = {k: v for k, v in diccionario_de_reemplazos.items() if k.startswith("<<fig_")}
    
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
                
                # Crear key de referencia: "<<fig_mapas>>" -> "<<ref_mapas>>"
                variable_ref_key = variable.replace("fig", "ref")
                
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


def insertar_referencias_cruzadas_en_plantilla(doc, diccionario_de_reemplazos: dict):
    """Reemplaza marcadores <<ref_*>> con referencias cruzadas a figuras.
    
    Args:
        doc: Objeto Document de python-docx
        diccionario_de_reemplazos: Diccionario retornado/mutado por insertar_figuras_en_plantilla
                                  Formato: {"<<ref_demo>>": ["_Ref_Fig_Mapa1", "_Ref_Fig_Temp"], ...}
    
    Comportamiento:
        - Busca variables que empiecen con "<<ref_" en los párrafos
        - Si la lista tiene 1 bookmark: inserta "Figura X"
        - Si la lista tiene 2+ bookmarks: inserta "Figura X a la Y"
        - X e Y son referencias cruzadas reales (campos REF) que muestran solo el número
        - Procesa TODOS los marcadores de un párrafo en una sola pasada
    
    Nota importante:
        Los bookmarks creados por crear_pie_de_figura solo incluyen el número de la figura,
        no el texto "Figura" ni el título. Por eso las referencias muestran solo el número.
    
    Ejemplo:
        Entrada en plantilla: "De la <<ref_demo_con_titulo>> se muestra el poder."
        Salida: "De la Figura 8 a la 10 se muestra el poder."
                (donde "8" y "10" son campos REF clickeables que muestran solo el número)
    """
    # Filtrar solo variables con bookmarks válidos
    variables_validas = {k: v for k, v in diccionario_de_reemplazos.items() 
                        if v is not None and k.startswith("<<ref_")}
    
    # Para cada párrafo
    for parrafo in doc.paragraphs:
        full_text = "".join(run.text for run in parrafo.runs)
        
        # Encontrar TODOS los marcadores <<ref_*>> en este párrafo
        marcadores_en_parrafo = []
        for variable_ref, lista_bookmarks in variables_validas.items():
            if variable_ref in full_text:
                # Encontrar todas las ocurrencias del marcador en el párrafo
                pos = full_text.find(variable_ref)
                if pos != -1:
                    marcadores_en_parrafo.append((pos, variable_ref, lista_bookmarks))
        
        # Si no hay marcadores en este párrafo, continuar al siguiente
        if not marcadores_en_parrafo:
            continue
        
        # Ordenar marcadores por posición (de izquierda a derecha)
        marcadores_en_parrafo.sort(key=lambda x: x[0])
        
        # Limpiar todos los runs del párrafo
        for run in parrafo.runs:
            run.text = ""
        
        # Reconstruir el párrafo procesando todos los marcadores
        pos_actual = 0
        
        for pos_marcador, variable_ref, lista_bookmarks in marcadores_en_parrafo:
            # Agregar texto antes del marcador
            if pos_marcador > pos_actual:
                texto_antes = full_text[pos_actual:pos_marcador]
                parrafo.add_run(texto_antes)
            
            # Insertar la referencia cruzada
            primer_bookmark = lista_bookmarks[0]
            ultimo_bookmark = lista_bookmarks[-1]
            
            if len(lista_bookmarks) == 1:
                # Caso: Solo una figura - "Figura X"
                aux_insertar_referencia_cruzada(parrafo, primer_bookmark, texto_antes="Figura", mostrar_numero=True)
            else:
                # Caso: Múltiples figuras - "Figura X a la Y"
                aux_insertar_referencia_cruzada(parrafo, primer_bookmark, texto_antes="Figura", mostrar_numero=True)
                parrafo.add_run(" a la ")
                aux_insertar_referencia_cruzada(parrafo, ultimo_bookmark, texto_antes="", mostrar_numero=True)
            
            # Avanzar posición actual
            pos_actual = pos_marcador + len(variable_ref)
        
        # Agregar texto después del último marcador
        if pos_actual < len(full_text):
            texto_despues = full_text[pos_actual:]
            parrafo.add_run(texto_despues)

    msg = "Referencias cruzadas insertadas."
    print(msg)


def reemplazar_texto_en_plantilla(doc, diccionario_de_reemplazos):
    """Reemplaza los marcadores de posición de texto en un documento de Word.
    
    Args:
        doc: El documento de Word (objeto Document de python-docx).
        diccionario_de_reemplazos: Diccionario donde las claves son los marcadores de posición 
                                  a buscar (por ejemplo, "<<orden_de_servicio>>") y los valores 
                                  son los textos que los reemplazarán (ej: "12345").
    
    Comportamiento:
        - Ignora marcadores de figuras ("<<fig_*>>"), referencias ("<<ref_*>>") y 
          documentos externos ("<<ruta_plan_de_crucero>>")
        - Procesa todas las variables de texto en cada párrafo de una sola vez
        - Preserva el formato del primer run del párrafo
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import reemplazar_texto_en_plantilla
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<orden_servicio>>": "12345",
            "<<cliente>>": "ACME Corporation",
            "<<fecha>>": "21/05/2026",
            "<<fig_mapa>>": [...],  # Se ignora aquí
        }
        
        reemplazar_texto_en_plantilla(doc, diccionario)
        doc.save('documento_con_texto.docx')
    """
    # Filtrar solo variables que NO son figuras, referencias, tablas o documentos externos
    variables_texto = {k: v for k, v in diccionario_de_reemplazos.items() 
                      if "fig" not in k and "ref" not in k and "tabla" not in k and "external_doc" not in k}
    
    # Para cada párrafo, procesar TODAS las variables de texto de una sola vez
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


def insertar_documento_externo_en_plantilla(doc, diccionario_de_reemplazos):
    """Inserta uno o más documentos Word externos en la plantilla.
    
    Args:
        doc: El documento de Word (objeto Document de python-docx).
        diccionario_de_reemplazos: Diccionario que debe contener la key "<<ruta_plan_de_crucero>>"
                                  con el valor siendo una ruta (string) o lista de rutas a 
                                  documentos Word externos.
    
    Comportamiento:
        - Busca el marcador "<<ruta_plan_de_crucero>>" en los párrafos
        - Inserta el contenido completo de cada documento externo (párrafos y tablas)
        - Copia imágenes y mantiene relaciones correctas
        - Excluye headers, footers, marcas de agua y configuraciones de sección
        - Agrega saltos de página entre múltiples documentos
    
    Ejemplo de uso:
        from docx import Document
        from word_template_writer import insertar_documento_externo_en_plantilla
        
        doc = Document('plantilla.docx')
        diccionario = {
            "<<ruta_plan_de_crucero>>": "plan_crucero.docx",
            # O múltiples documentos:
            # "<<ruta_plan_de_crucero>>": ["plan1.docx", "plan2.docx"],
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
    """Rellena una tabla en la plantilla con datos de un DataFrame o diccionario.
    
    Soporta estilos configurables, colores de fondo, MultiIndex, y merge de celdas.
    
    Args:
        doc: Objeto Document de python-docx
        nombre_marcador: Marcador a buscar en la tabla (ej: "<<table_sondas>>")
        datos: DataFrame o diccionario con los datos a insertar
              Formato dict: {"col1": [val1, val2, ...], "col2": [...], ...}
        config_estilos: Configuración de estilos (EstilosTabla, dict, o None)
                       Si None, usa estilos por defecto razonables
        opciones_tabla: Opciones de tabla (OpcionesTabla, dict, o None)
                       Si None, usa opciones por defecto
    
    Comportamiento:
        - Busca el marcador en cualquier celda de las tablas del documento
        - Inserta datos del DataFrame fila por fila, respetando orden de columnas
        - Aplica estilos según jerarquía: celda > fila > columna > defecto
        - Soporta colores de fondo RGB y estilos de párrafo del documento
        - Aplana MultiIndex automáticamente si está activado en opciones
        - Detecta y combina celdas verticales con valores repetidos (si está activado)
        - Elimina fila marcador automáticamente (si está activado)
    
    Ejemplo de uso básico (sin estilos personalizados):
        >>> from docx import Document
        >>> doc = Document('plantilla.docx')
        >>> datos = {
        ...     'Secuencia': [0, 1, 2],
        ...     'Localización': ['BOT-01', 'BOT-02', 'BOT-03'],
        ...     'Latitud': [18.5, 18.6, 18.7]
        ... }
        >>> rellenar_tablas_en_plantilla(doc, "<<table_posiciones>>", datos)
        >>> doc.save('documento_con_tabla.docx')
    
    Ejemplo avanzado (con estilos personalizados):
        >>> from docx import Document
        >>> from crear_documentos.services.word_template_writer import EstilosTabla, OpcionesTabla
        >>> 
        >>> doc = Document('plantilla.docx')
        >>> 
        >>> # Configurar estilos
        >>> estilos = EstilosTabla(doc)
        >>> estilos.set_color_de_columna(0, (230, 230, 250))  # Lavanda para primera columna
        >>> estilos.set_estilo_de_columna(1, 'texto_tablas_justificado')
        >>> estilos.set_color_de_fila(0, (255, 240, 245))  # Rosado para primera fila
        >>> 
        >>> # Configurar opciones
        >>> opciones = OpcionesTabla()
        >>> opciones.set_detectar_merge(True)
        >>> opciones.set_columnas_para_merge([0, 1])  # Solo merge en columnas 0 y 1
        >>> 
        >>> # DataFrame con datos
        >>> import pandas as pd
        >>> df = pd.DataFrame({
        ...     'Tipo': ['A', 'A', 'B', 'B'],
        ...     'Subtipo': ['A1', 'A1', 'B1', 'B2'],
        ...     'Valor': [10, 20, 30, 40]
        ... })
        >>> 
        >>> rellenar_tablas_en_plantilla(doc, "<<table_datos>>", df, estilos, opciones)
        >>> doc.save('documento_avanzado.docx')
    
    Ejemplo con diccionario de configuración (backward compatibility):
        >>> config_dict = {
        ...     "por_defecto": {"estilo_parrafo": "Normal", "altura_fila": 288290},
        ...     "por_columna": {0: {"color_fondo": (230, 230, 250)}},
        ...     "por_fila": {},
        ...     "por_celda": {}
        ... }
        >>> opciones_dict = {
        ...     "aplanar_multiindex": True,
        ...     "detectar_merge": False,
        ...     "columnas_para_merge": None,
        ...     "eliminar_fila_marcador": True
        ... }
        >>> rellenar_tablas_en_plantilla(doc, "<<table_datos>>", datos, config_dict, opciones_dict)
    
    Raises:
        ValueError: Si el marcador no se encuentra, si hay errores en la configuración,
                   o si los datos están vacíos
    
    Nota:
        - La plantilla debe tener un marcador <<nombre>> en alguna celda de la tabla
        - El marcador se coloca típicamente en la primera fila de datos (no encabezado)
        - Los datos se insertan por ORDEN de columna, no por nombre
        - Columna 0 del DataFrame → Columna 0 de la tabla
    """
    
    variables_texto = {k: v for k, v in diccionario_de_reemplazos.items() if k.startswith("<<tabla_")}
    
    for variable in variables_texto.keys():
        df_tabla = diccionario_de_reemplazos[variable]["tabla"]
        config_estilos = diccionario_de_reemplazos[variable]["estilos_de_tabla"]
        opciones_tabla = diccionario_de_reemplazos[variable]["opciones_de_tabla"]
        fill_table(doc, variable, df_tabla, config_estilos, opciones_tabla)
    
        msg = f"Tabla '{variable}' rellenada correctamente."
        print(msg)


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
        - Ignora variables que inician con: <<fig_, <<ref_, <<tabla_, <<external_doc
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
    replace_text_variables_in_tables(doc, diccionario_de_reemplazos)

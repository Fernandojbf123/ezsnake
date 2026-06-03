"""
Internal helper functions for table manipulation in Word templates.

This module contains low-level functions for filling Word tables with DataFrame data,
including support for MultiIndex DataFrames, merged cells, and flexible styling.

Private module - Not intended for direct external use.
Import from the public API in api.py instead.
"""

import pandas as pd
from typing import Union, Optional, Dict, List, Tuple
from docx.oxml.ns import qn
from .schemas_helpers import EstilosTabla, OpcionesTabla
from ._table_styling import apply_cell_style, apply_row_height


def insertar_titulo_de_tabla_con_bookmark(doc, marcador_tabla, titulo, bookmark, estilo_titulo="Normal"):
    """Inserta un título con bookmark antes de la tabla que contiene el marcador."""
    if not titulo:
        return

    tabla_objetivo = None
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if marcador_tabla in celda.text:
                    tabla_objetivo = tabla
                    break
            if tabla_objetivo is not None:
                break
        if tabla_objetivo is not None:
            break

    if tabla_objetivo is None:
        raise ValueError(f"No se encontró la tabla con marcador '{marcador_tabla}' para insertar título.")

    from docx.oxml import OxmlElement

    bookmark_id = str(abs(hash(bookmark)) % 1000000)

    nuevo_p = OxmlElement('w:p')

    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), estilo_titulo)
    pPr.append(pStyle)
    nuevo_p.append(pPr)

    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), bookmark_id)
    bookmark_start.set(qn('w:name'), bookmark)
    nuevo_p.append(bookmark_start)

    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.set(qn('xml:space'), 'preserve')
    text.text = str(titulo)
    run.append(text)
    nuevo_p.append(run)

    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), bookmark_id)
    nuevo_p.append(bookmark_end)

    tabla_objetivo._element.addprevious(nuevo_p)


# ===== NORMALIZATION FUNCTIONS =====

def _normalize_input_to_dataframe(datos: Union[pd.DataFrame, dict]) -> pd.DataFrame:
    """
    Normaliza la entrada a DataFrame.
    
    Args:
        datos: DataFrame o diccionario de columnas
    
    Returns:
        DataFrame normalizado
    
    Raises:
        ValueError: Si los datos están vacíos o no son válidos
    """
    if isinstance(datos, pd.DataFrame):
        df = datos
    elif isinstance(datos, dict):
        df = pd.DataFrame(datos)
    else:
        raise ValueError(
            f"datos debe ser DataFrame o diccionario. Recibido: {type(datos)}"
        )
    
    if df.empty:
        raise ValueError("El DataFrame está vacío. No hay datos para insertar.")
    
    return df


def _normalize_config_estilos(config_estilos: Union[EstilosTabla, dict, None], doc) -> dict:
    """
    Normaliza la configuración de estilos a diccionario.
    
    Args:
        config_estilos: Objeto EstilosTabla, dict, o None
        doc: Objeto Document (para crear EstilosTabla por defecto si es None)
    
    Returns:
        Diccionario de configuración normalizado
    """
    if config_estilos is None:
        # Crear EstilosTabla por defecto
        config_obj = EstilosTabla(doc)
        return config_obj.to_dict()
    elif isinstance(config_estilos, EstilosTabla):
        return config_estilos.to_dict()
    elif isinstance(config_estilos, dict):
        # Validar estructura básica del diccionario
        required_keys = ["por_defecto", "por_columna", "por_fila", "por_celda"]
        if not all(key in config_estilos for key in required_keys):
            raise ValueError(
                f"Diccionario config_estilos debe contener las claves: {required_keys}"
            )
        return config_estilos
    else:
        raise ValueError(
            f"config_estilos debe ser EstilosTabla, dict o None. Recibido: {type(config_estilos)}"
        )


def _normalize_opciones_tabla(opciones_tabla: Union[OpcionesTabla, dict, None]) -> dict:
    """
    Normaliza las opciones de tabla a diccionario.
    
    Args:
        opciones_tabla: Objeto OpcionesTabla, dict, o None
    
    Returns:
        Diccionario de opciones normalizado
    """
    if opciones_tabla is None:
        # Crear OpcionesTabla por defecto
        opciones_obj = OpcionesTabla()
        return opciones_obj.to_dict()
    elif isinstance(opciones_tabla, OpcionesTabla):
        return opciones_tabla.to_dict()
    elif isinstance(opciones_tabla, dict):
        # Agregar valores por defecto para claves faltantes
        defaults = OpcionesTabla().to_dict()
        return {**defaults, **opciones_tabla}
    else:
        raise ValueError(
            f"opciones_tabla debe ser OpcionesTabla, dict o None. Recibido: {type(opciones_tabla)}"
        )


# ===== TABLE SEARCH FUNCTIONS =====

def _find_table_with_marker(doc, nombre_marcador: str) -> Tuple[object, int, int]:
    """
    Busca una tabla que contenga el marcador especificado.
    
    Args:
        doc: Objeto Document
        nombre_marcador: Marcador a buscar (ej: "<<table_sondas>>")
    
    Returns:
        Tupla (tabla, fila_marcador, columna_marcador)
    
    Raises:
        ValueError: Si no se encuentra el marcador
    """
    for table in doc.tables:
        # Buscar en todas las celdas de la tabla
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                if nombre_marcador in cell.text:
                    return (table, row_idx, col_idx)
    
    raise ValueError(
        f"No se encontró el marcador '{nombre_marcador}' en ninguna tabla del documento."
    )


def _disable_row_header_repeat(row):
    """
    Desactiva la propiedad de repetición de encabezado en una fila.
    
    En Word, las filas de encabezado tienen la propiedad tblHeader que hace
    que se repitan en cada salto de página. Esta función elimina esa propiedad.
    
    Args:
        row: Objeto Row de python-docx
    """
    tr = row._element
    trPr = tr.get_or_add_trPr()
    
    # Buscar y eliminar el elemento tblHeader si existe
    tblHeader = trPr.find(qn('w:tblHeader'))
    if tblHeader is not None:
        trPr.remove(tblHeader)


def _delete_row(table, row_idx: int):
    """
    Elimina una fila de la tabla manipulando el XML directamente.
    
    python-docx no proporciona una API directa para eliminar filas,
    por lo que se debe manipular el XML.
    
    Args:
        table: Objeto Table de python-docx
        row_idx: Índice de la fila a eliminar (0-based)
    """
    tbl = table._element
    tr = table.rows[row_idx]._element
    tbl.remove(tr)


# ===== STYLE MERGING FUNCTIONS =====

def _merge_style_config(config: dict, row_idx: int, col_idx: int) -> dict:
    """
    Combina la configuración de estilos para una celda específica.
    
    Resuelve la jerarquía: celda > fila > columna > defecto
    
    Args:
        config: Diccionario de configuración completo
        row_idx: Índice de fila (0-based)
        col_idx: Índice de columna (0-based)
    
    Returns:
        Diccionario con configuración mergeada para esta celda
    """
    merged = {}
    
    # 1. Aplicar configuración por defecto
    if "por_defecto" in config:
        merged.update(config["por_defecto"])
    
    # 2. Aplicar configuración de columna (sobrescribe defecto)
    if "por_columna" in config and col_idx in config["por_columna"]:
        merged.update(config["por_columna"][col_idx])
    
    # 3. Aplicar configuración de fila (sobrescribe columna y defecto)
    if "por_fila" in config and row_idx in config["por_fila"]:
        merged.update(config["por_fila"][row_idx])
    
    # 4. Aplicar configuración de celda (sobrescribe todo)
    if "por_celda" in config and (row_idx, col_idx) in config["por_celda"]:
        merged.update(config["por_celda"][(row_idx, col_idx)])
    
    return merged


# ===== MULTIINDEX FLATTENING FUNCTIONS =====

def _flatten_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplana un DataFrame con MultiIndex en filas o columnas.
    
    Args:
        df: DataFrame potencialmente con MultiIndex
    
    Returns:
        DataFrame aplanado con índice simple
    """
    df_flat = df.copy()
    
    # Aplanar MultiIndex en filas (índice)
    if isinstance(df_flat.index, pd.MultiIndex):
        df_flat = df_flat.reset_index()
    
    # Aplanar MultiIndex en columnas
    if isinstance(df_flat.columns, pd.MultiIndex):
        # Unir niveles con '_'
        df_flat.columns = ['_'.join(map(str, col)).strip('_') for col in df_flat.columns]
    
    return df_flat


# ===== MERGED CELLS FUNCTIONS =====

def _detect_merge_regions_vertical(
    df: pd.DataFrame,
    columnas_para_merge: Optional[List[int]]
) -> List[Tuple[int, int, int]]:
    """
    Detecta regiones de celdas a combinar verticalmente.
    
    Busca valores consecutivos repetidos en columnas especificadas.
    
    Args:
        df: DataFrame con los datos
        columnas_para_merge: Lista de índices de columnas o None (todas)
    
    Returns:
        Lista de tuplas (fila_inicio, fila_fin, columna)
    """
    regions = []
    
    # Determinar qué columnas analizar
    if columnas_para_merge is None:
        cols_to_check = range(len(df.columns))
    else:
        cols_to_check = columnas_para_merge
    
    for col_idx in cols_to_check:
        if col_idx >= len(df.columns):
            continue  # Saltar si el índice está fuera de rango
        
        col_name = df.columns[col_idx]
        values = df[col_name].tolist()
        
        # Detectar secuencias de valores repetidos
        start_idx = 0
        while start_idx < len(values):
            current_value = values[start_idx]
            end_idx = start_idx
            
            # Contar cuántas filas consecutivas tienen el mismo valor
            while end_idx + 1 < len(values) and values[end_idx + 1] == current_value:
                end_idx += 1
            
            # Si hay al menos 2 celdas consecutivas con el mismo valor, agregar región
            if end_idx > start_idx:
                regions.append((start_idx, end_idx, col_idx))
            
            start_idx = end_idx + 1
    
    return regions


def _apply_merged_cells_vertical(table, regions: List[Tuple[int, int, int]], fila_inicio_datos: int):
    """
    Aplica merged cells verticales a la tabla.
    
    Args:
        table: Objeto Table de python-docx
        regions: Lista de tuplas (fila_inicio, fila_fin, columna)
        fila_inicio_datos: Índice de la fila donde inician los datos (después del marcador)
    """
    for start_row, end_row, col in regions:
        # Ajustar índices de fila para la tabla real (sumar fila_inicio_datos)
        table_start_row = fila_inicio_datos + start_row
        table_end_row = fila_inicio_datos + end_row
        
        try:
            # Obtener celdas a combinar
            cell_start = table.cell(table_start_row, col)
            cell_end = table.cell(table_end_row, col)
            
            # Combinar celdas
            cell_start.merge(cell_end)
        except Exception as e:
            # Si falla el merge, continuar con las demás (no es crítico)
            print(f"Advertencia: No se pudo combinar celdas ({start_row}-{end_row}, {col}): {e}")


# ===== MAIN FILL TABLE FUNCTION =====

def fill_table(
    doc,
    nombre_marcador: str,
    datos: Union[pd.DataFrame, dict],
    config_estilos: Union[EstilosTabla, dict, None] = None,
    opciones_tabla: Union[OpcionesTabla, dict, None] = None
):
    """
    Rellena una tabla dinámicamente buscando un marcador e insertando datos de un DataFrame.
    
    Esta es la función principal refactorizada que orquesta todo el proceso.
    
    Args:
        doc: Objeto Document de python-docx
        nombre_marcador: Marcador a buscar en la tabla (ej: "<<table_sondas>>")
        datos: DataFrame o diccionario con los datos a insertar
        config_estilos: Configuración de estilos (objeto EstilosTabla, dict, o None)
        opciones_tabla: Opciones de tabla (objeto OpcionesTabla, dict, o None)
    
    Returns:
        Objeto Document modificado
    
    Raises:
        ValueError: Si hay errores en la configuración o datos
    
    Flujo:
        1. Normalizar inputs (DataFrame, config, opciones)
        2. Buscar tabla con marcador
        3. Aplanar DataFrame si tiene MultiIndex
        4. Insertar datos fila por fila
        5. Detectar y aplicar merged cells (si está activado)
        6. Aplicar estilos a cada celda
        7. Eliminar fila marcador (si está activado)
    """
    # ===== PASO 1: NORMALIZAR INPUTS =====
    df = _normalize_input_to_dataframe(datos)
    config = _normalize_config_estilos(config_estilos, doc)
    opciones = _normalize_opciones_tabla(opciones_tabla)
    
    # ===== PASO 2: BUSCAR TABLA CON MARCADOR =====
    table, fila_marcador, col_marcador = _find_table_with_marker(doc, nombre_marcador)
    
    # ===== BUG FIX 1: DESACTIVAR REPETICIÓN DE ENCABEZADO EN FILA MARCADOR =====
    # La fila del marcador puede tener la propiedad tblHeader (encabezado repetible)
    # que hace que se repita en cada salto de página. Debemos desactivarla.
    _disable_row_header_repeat(table.rows[fila_marcador])
    
    # ===== PASO 3: APLANAR DATAFRAME SI TIENE MULTIINDEX =====
    if opciones["aplanar_multiindex"]:
        if isinstance(df.index, pd.MultiIndex) or isinstance(df.columns, pd.MultiIndex):
            df = _flatten_dataframe(df)
    else:
        # Si no se debe aplanar y tiene MultiIndex, lanzar error
        if isinstance(df.index, pd.MultiIndex) or isinstance(df.columns, pd.MultiIndex):
            raise ValueError(
                "El DataFrame tiene MultiIndex pero 'aplanar_multiindex' está desactivado. "
                "Active la opción o aplane el DataFrame manualmente con df.reset_index()."
            )
    
    # ===== PASO 4: VALIDAR QUE HAY SUFICIENTES COLUMNAS =====
    num_cols_df = len(df.columns)
    num_cols_table = len(table.columns)
    if num_cols_df > num_cols_table:
        raise ValueError(
            f"El DataFrame tiene {num_cols_df} columnas pero la tabla solo tiene {num_cols_table}. "
            "Ajuste la plantilla o el DataFrame."
        )
    
    # ===== PASO 5: AGREGAR FILAS NECESARIAS =====
    # Calcular cuántas filas necesitamos agregar
    # Si eliminamos el marcador: usamos esa fila + las que siguen
    # Si NO eliminamos: necesitamos todas las filas de datos desde el marcador
    filas_necesarias = len(df)
    filas_actuales_disponibles = len(table.rows) - fila_marcador  # Incluye fila marcador
    
    if filas_necesarias > filas_actuales_disponibles:
        # Agregar solo las filas faltantes (BUG FIX 2: cálculo corregido)
        filas_a_agregar = filas_necesarias - filas_actuales_disponibles
        for _ in range(filas_a_agregar):
            table.add_row()
    
    # ===== PASO 6: DETECTAR REGIONES DE MERGE (antes de insertar datos) =====
    # BUG FIX 3: Detectar merge ANTES de insertar texto para evitar duplicados
    merge_regions = []
    if opciones["detectar_merge"]:
        merge_regions = _detect_merge_regions_vertical(df, opciones["columnas_para_merge"])
    
    # Crear un set de celdas que NO deben recibir texto (segunda celda en adelante de merge)
    celdas_merged_secundarias = set()
    for start_row, end_row, col in merge_regions:
        # Agregar todas las celdas excepto la primera de cada región
        for row_idx in range(start_row + 1, end_row + 1):
            celdas_merged_secundarias.add((row_idx, col))
    
    # ===== PASO 7: INSERTAR DATOS FILA POR FILA =====
    # Siempre empezamos a escribir desde la fila del marcador
    fila_inicio_datos = fila_marcador
    
    for df_row_idx, (_, df_row) in enumerate(df.iterrows()):
        # Índice de fila en la tabla (después del marcador)
        table_row_idx = fila_inicio_datos + df_row_idx
        
        # Insertar datos en cada columna
        for df_col_idx, col_name in enumerate(df.columns):
            # BUG FIX 3: No escribir en celdas secundarias de regiones mergeadas
            if (df_row_idx, df_col_idx) not in celdas_merged_secundarias:
                cell = table.cell(table_row_idx, df_col_idx)
                cell.text = str(df_row[col_name])
    
    # ===== PASO 8: APLICAR MERGED CELLS =====
    if opciones["detectar_merge"]:
        _apply_merged_cells_vertical(table, merge_regions, fila_inicio_datos)
    
    # ===== PASO 9: APLICAR ESTILOS A CADA CELDA =====
    for df_row_idx in range(len(df)):
        table_row_idx = fila_inicio_datos + df_row_idx
        row = table.rows[table_row_idx]
        
        # Aplicar altura de fila
        altura_fila = config["por_defecto"].get("altura_fila")
        if altura_fila:
            apply_row_height(row, altura_fila)
        
        # Aplicar estilos a cada celda
        # BUG FIX 3: Solo aplicar estilos a celdas que no son secundarias en merge
        for df_col_idx in range(len(df.columns)):
            if (df_row_idx, df_col_idx) not in celdas_merged_secundarias:
                cell = table.cell(table_row_idx, df_col_idx)
                merged_config = _merge_style_config(config, df_row_idx, df_col_idx)
                apply_cell_style(cell, doc, config, df_row_idx, df_col_idx, merged_config)
    
    # ===== PASO 10: ELIMINAR FILAS SOBRANTES AL FINAL =====
    # BUG FIX 2: Eliminar filas vacías que puedan quedar al final de la tabla
    # Esto puede ocurrir si la plantilla tiene filas preexistentes o si add_row()
    # creó filas de más
    ultima_fila_con_datos = fila_inicio_datos + len(df) - 1
    total_filas = len(table.rows)
    
    # Eliminar todas las filas después de la última fila con datos
    # Nota: Eliminamos de atrás hacia adelante para evitar problemas de índices
    for row_idx in range(total_filas - 1, ultima_fila_con_datos, -1):
        _delete_row(table, row_idx)
    
    return doc


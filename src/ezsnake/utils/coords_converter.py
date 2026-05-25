import re

def convertir_cualquier_coordenada_a_grados_decimales(coordenada: str | float) -> float:
    """Convierte coordenadas geográficas de formato DMS o DM a grados decimales.
    
    Parámetros:
        coordenada (str | float): Coordenada en formato DMS, DM o decimal
            Ejemplos: 18° 44' 20" N, 18° 44.53630' N, 18.738889
    
    Retorna:
        float: Coordenada en grados decimales
    
    Ejemplos:
        >>> convertir_cualquier_coordenada_a_grados_decimales("18° 44' 20\" N")
        18.738889
        >>> convertir_cualquier_coordenada_a_grados_decimales("18° 44.53630' N")
        18.742272
        >>> convertir_cualquier_coordenada_a_grados_decimales(18.738889)
        18.738889
        
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    # Si ya es un número (float o int), retornarlo directamente
    try:
        return float(coordenada)
    except:
        pass
    
    # Eliminar espacios extras
    coordenada = coordenada.strip()
    
    # Extraer la dirección (N, S, E, W)
    direccion = coordenada[-1].upper()
    
    # Patron para DMS: grados° minutos' segundos'' dirección
    patron_dms = r"(\d+)°\s*(\d+)'\s*(\d+(?:\.\d+)?)\"\s*[NSEW]"
    # Patron para DM: grados° minutos.decimales' dirección
    patron_dm = r"(\d+)°\s*(\d+(?:\.\d+)?)'\s*[NSEW]"
    
    match_dms = re.match(patron_dms, coordenada)
    match_dm = re.match(patron_dm, coordenada)
    
    if match_dms:
        # Formato DMS (grados, minutos, segundos)
        grados = float(match_dms.group(1))
        minutos = float(match_dms.group(2))
        segundos = float(match_dms.group(3))
        decimal = grados + (minutos / 60) + (segundos / 3600)
    elif match_dm:
        # Formato DM (grados, minutos decimales)
        grados = float(match_dm.group(1))
        minutos = float(match_dm.group(2))
        decimal = grados + (minutos / 60)
    else:
        raise ValueError(f"Formato de coordenada no válido: {coordenada}")
    
    # Aplicar signo según la dirección
    if direccion in ['S', 'W']:
        decimal = -decimal
    
    return round(decimal,6)


def convertir_cualquier_coordenada_a_grados_y_minutos(coordenada: str | float, tipo: str = 'lat') -> str:
    """Convierte coordenadas geográficas de cualquier formato a grados y minutos.
    
    Parámetros:
        coordenada (str | float): Coordenada en formato DMS, DM o decimal
            Ejemplos: 18.5792, "18° 44.53630' N", "18° 44' 20\" N"
        tipo (str): Tipo de coordenada - 'lat' para latitud o 'lon' para longitud
            Por defecto 'lat'. Determina si usar N/S (lat) o E/W (lon)
    
    Retorna:
        str: Coordenada en formato grados y minutos (e.g., "18° 44.53630' N")
    
    Ejemplos:
        >>> convertir_cualquier_coordenada_a_grados_y_minutos(18.5792)
        "18° 34.75200' N"
        >>> convertir_cualquier_coordenada_a_grados_y_minutos(-92.5487, 'lon')
        "92° 32.92200' W"
        >>> convertir_cualquier_coordenada_a_grados_y_minutos("18° 44' 20\" N")
        "18° 44.33333' N"
        >>> convertir_cualquier_coordenada_a_grados_y_minutos("18° 44.53630' N")
        "18° 44.53630' N"
    """
    # Si es una cadena, verificar si ya está en formato DM
    if isinstance(coordenada, str):
        coordenada_str = coordenada.strip()
        
        # Patron para DM: grados° minutos.decimales' dirección
        patron_dm = r"(\d+)°\s*(\d+(?:\.\d+)?)'\s*[NSEW]"
        match_dm = re.match(patron_dm, coordenada_str)
        
        if match_dm:
            # Ya está en el formato correcto, retornar normalizando el formato
            grados = match_dm.group(1)
            minutos = match_dm.group(2)
            direccion = coordenada_str[-1].upper()
            return f"{grados}° {minutos}' {direccion}"
        
        # Si está en DMS u otro formato, convertir primero a decimal
        decimal = convertir_cualquier_coordenada_a_grados_decimales(coordenada)
    else:
        # Es un valor numérico (float o int)
        decimal = float(coordenada)
    
    # Convertir de grados decimales a grados y minutos
    # Determinar la dirección
    if tipo.lower() == 'lat':
        direccion = 'N' if decimal >= 0 else 'S'
    else:  # lon
        direccion = 'E' if decimal >= 0 else 'W'
    
    # Trabajar con el valor absoluto
    decimal_abs = abs(decimal)
    
    # Extraer grados (parte entera)
    grados = int(decimal_abs)
    
    # Calcular minutos (parte decimal * 60)
    minutos = (decimal_abs - grados) * 60
    
    # Formatear el resultado
    return f"{grados}° {minutos:.5f}' {direccion}"


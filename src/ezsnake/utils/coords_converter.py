import re

def convertir_cualquier_coordenada_a_grados_decimales(coordenada: str | float) -> float:
    """Convierte coordenadas geográficas de formato DMS o DM a grados decimales.
    
    Parámetros:
        coordenada (str | float): Coordenada en formato DMS, DM o decimal
            Ejemplos: "18° 44' 20'' N", "18° 44.5363' N", 18.738889
    
    Retorna:
        float: Coordenada en grados decimales
    
    Ejemplos:
        >>> convertir_cualquier_coordenada_a_grados_decimales("18° 44' 20'' N")
        18.738889
        >>> convertir_cualquier_coordenada_a_grados_decimales("18° 44.5363' N")
        18.742272
        >>> convertir_cualquier_coordenada_a_grados_decimales(18.738889)
        18.738889
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
    patron_dms = r"(\d+)°\s*(\d+)'\s*(\d+(?:\.\d+)?)''\s*[NSEW]"
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



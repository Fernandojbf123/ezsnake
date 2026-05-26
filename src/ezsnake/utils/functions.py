import numpy as np
import pandas as pd

def uv2polar(u: float | np.ndarray, v: float | np.ndarray) -> tuple[float | np.ndarray, float | np.ndarray]:
    """"
    Convierte componentes u (este) y v (norte) a dirección en grados y magnitud.
    Entradas:
    u: Componente este (float o array)
    v: Componente norte (float o array)
    Salidas:
    dir_deg: Dirección en grados (float o array)
    spd: Magnitud o velocidad (float o array)
    
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    # Magnitud
    spd = np.hypot(u, v)
    # Dirección en radianes (como MATLAB: cart2pol(v, u))
    dir_rad = np.arctan2(u, v)
    # Ajuste de ángulo negativo
    dir_rad = np.where(dir_rad < 0, dir_rad + 2 * np.pi, dir_rad)
    # Convertir a grados
    dir_deg = np.degrees(dir_rad)
    return dir_deg, spd


def polar2uv(dir_deg: float | np.ndarray, spd: float | np.ndarray) -> tuple[float | np.ndarray, float | np.ndarray]:
    """" 
    Convierte dirección en grados y magnitud a componentes u (este) y v (norte).
    Entradas:
    dir_deg: Dirección en grados (float o array)
    spd: Magnitud o velocidad (float o array)
    Salidas:
    u: Componente este (float o array)
    v: Componente norte (float o array)
    
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    # Convertir grados a radianes
    dir_rad = np.radians(dir_deg)
    # Componente este (u)
    u = spd * np.sin(dir_rad)
    # Componente norte (v)
    v = spd * np.cos(dir_rad)
    return u, v


def grados_a_km_lat(delta_lat: float) -> float:
    """
    Calcula la distancia en kilómetros correspondiente a un cambio de latitud (delta_lat).
    Entrada:
    delta_lat: cambio en latitud en grados decimales.
    Salida:
    distancia en kilómetros correspondiente al cambio de latitud.
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********    
    """
    delta_lat = abs(delta_lat)
    return delta_lat * 111.32


def grados_a_km_lon(delta_lon: float, latitud: float) -> float:
    """
    Calcula la distancia en kilómetros correspondiente a un cambio de longitud (delta_lon).
    Entradas:
    delta_lon: cambio en longitud en grados decimales
    latitud: latitud en grados decimales
    Salida:
    distancia en kilómetros correspondiente al cambio de longitud.
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    delta_lon = abs(delta_lon)
    latitud = abs(latitud)
    resultado = delta_lon * 111.32 * np.cos(np.radians(latitud))
    if abs(resultado) < 1e-6:  # Evitar valores extremadamente pequeños
        resultado = 0.0
    return resultado


def distancia_entre_dos_puntos(lon2: float, lat2:float, lon1: float, lat1: float, unidad='km') -> float:
    """"
    Calcula la distancia entre dos coordenadas geográficas 
    Entradas: 
    lon1 = Float, longitud inicial en grados decimales
    lat1 = Float, latitud inicial en grados decimales
    lon2 = Float, longitud final en grados decimales
    lat2 = Float, latitud final en grados decimales
    unidad = String, 'km' para kilómetros o 'mn' para millas náuticas
    
    Salida:
    distancia = Float, distancia entre las coordenadas en la unidad especificada
    
    Notas: 1 km = 1.852 millas náuticas
    
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    delta_lat = abs(lat2 - lat1)
    delta_lon = abs(lon2 - lon1)
    
    delta_lat = grados_a_km_lat(delta_lat)
    delta_lon = grados_a_km_lon(delta_lon, (lat1 + lat2) / 2)
    distancia = np.sqrt(delta_lon**2 + delta_lat**2)    
    if unidad == 'mn':
        distancia = distancia / 1.852  # Convertir kilómetros a millas náuticas
    return distancia


def calcular_tiempo_de_viaje(lon2: float, lat2: float, lon1: float, lat1: float, speed: float, unidad='km') -> float:
    """ Entradas: 
    lon1 = Float, longitud inicial en grados decimales
    lat1 = Float, latitud inicial en grados decimales
    lon2 = Float, longitud final en grados decimales
    lat2 = Float, latitud final en grados decimales
    speed = Float, velocidad de viaje en km/h o nudos (dependiendo de la unidad)
    unidad = String, 'km' para kilómetros o 'mn' para millas náuticas  
    
    Salida:
    tiempo = Float, tiempo de viaje entre las coordenadas en la unidad especificada
    
    Notas: 1 km = 1.852 millas náuticas
    
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    speed = abs(speed)
    
    distancia = distancia_entre_dos_puntos(lon2, lat2, lon1, lat1, unidad)
    if unidad == 'mn':
        # Convertir velocidad de nudos a km/h (1 nudo = 1.852 km/h)
        speed = speed * 1.852

    # Calcular tiempo de viaje en horas
    tiempo_horas = distancia / speed
    return tiempo_horas


def timestamp_a_texto_espanol(fecha: pd.Timestamp, mes_y_anio: bool) -> str:
    """
    Convierte un timestamp de pandas a un formato de texto en español.
    Entradas:
    fecha: timestamp de pandas
    mes_y_anio: bool, True para mostrar solo mes y año, False para mostrar día, mes y año
    Salida:
    texto en formato español
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    meses = {
        1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
        5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
        9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
    }
    
    dia = fecha.day
    mes = meses[fecha.month]
    anio = fecha.year
    
    if mes_y_anio:
        return f"{mes} de {anio}"
    
    return f"{dia:02d} de {mes} de {anio}"


def get_excel_variables_name(df_datos_documento: pd.DataFrame) -> list:
    """ 
    Busca en la columna 0 de un dataframe de pandas; devuelve una lista con los nombres de las variables.
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    varnames = []
    varnames = df_datos_documento.iloc[:,0].notna().tolist()
    return varnames


def get_excel_variable_values(df_datos_documento: pd.DataFrame, nombre_variable: str) -> None | np.ndarray:
    """ 
    Lee la columna 0 de un dataframe de pandas; consigue la fila con el nombre de variable; 
    devuelve el resto de la fila como un array de numpy, sin valores NaN.
    
    Esto es algo típico en un excel que contiene datos para una plantilla de excel.
    
    Entradas:
    df_datos_documento: DataFrame de pandas con los datos del documento Excel
    nombre_variable: String con el nombre de la variable a buscar en la columna 0
    Salidas:
    varvalue: Array de numpy con los valores de la variable, sin NaN; o None si no se encuentra la variable
    
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """

    varvalue = None
    for row in df_datos_documento.iterrows():
        if row[1][0] == nombre_variable:
            varvalue = np.array(row[1][1:])
            varvalue = varvalue[~pd.isna(varvalue)]  # Eliminar valores NaN
            return varvalue
    
    if varvalue is None:
        raise ValueError(f"No se encontró la variable '{nombre_variable}' en el DataFrame de datos del documento.")


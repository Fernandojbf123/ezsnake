import numpy as np

def uv2polar(u, v):
    # Magnitud
    spd = np.hypot(u, v)
    # Dirección en radianes (como MATLAB: cart2pol(v, u))
    dir_rad = np.arctan2(u, v)
    # Ajuste de ángulo negativo
    dir_rad = np.where(dir_rad < 0, dir_rad + 2 * np.pi, dir_rad)
    # Convertir a grados
    dir_deg = np.degrees(dir_rad)
    return dir_deg, spd

def polar2uv(dir_deg, spd):
    # Convertir grados a radianes
    dir_rad = np.radians(dir_deg)
    # Componente este (u)
    u = spd * np.sin(dir_rad)
    # Componente norte (v)
    v = spd * np.cos(dir_rad)
    return u, v



def grados_a_km_lat(delta_lat):
    """Calcula la distancia en kilómetros correspondiente a un cambio de latitud (delta_lat).
        delta_lat: cambio en latitud en grados decimales
    """
    if abs(delta_lat) < 1e-6:  # Evitar valores extremadamente pequeños
        raise ValueError("La diferencia entre latitudes (delta_lat) debe ser mayor que cero.")
    return delta_lat * 111.32


def grados_a_km_lon(delta_lon, latitud):
    """Calcula la distancia en kilómetros correspondiente a un cambio de longitud (delta_lon).
        delta_lon: cambio en longitud en grados decimales
    """
    if delta_lon < 1e-6 or latitud < 1e-6:  # Evitar valores extremadamente pequeños
        raise ValueError("La diferencia entre longitudes (delta_lon) y la coordenada de latitud deben ser mayores que cero.")
    resultado = delta_lon * 111.32 * np.cos(np.radians(latitud))
    if abs(resultado) < 1e-6:  # Evitar valores extremadamente pequeños
        resultado = 0.0
    return resultado


def distancia_entre_dos_puntos(lon_fin, lat_fin, lon_ini, lat_ini, unidad='km'):
    """" Calcula la distancia entre dos coordenadas geográficas 
    Entradas: 
    lon_ini = Float, longitud inicial en grados decimales
    lat_ini = Float, latitud inicial en grados decimales
    lon_fin = Float, longitud final en grados decimales
    lat_fin = Float, latitud final en grados decimales
    unidad = String, 'km' para kilómetros o 'mn' para millas náuticas
    
    Salida:
    distancia = Float, distancia entre las coordenadas en la unidad especificada
    
    Notas: 1 km = 1.852 millas náuticas
    """
    delta_lat = grados_a_km_lat(lat_fin - lat_ini)
    delta_lon = grados_a_km_lon(lon_fin - lon_ini, (lat_ini + lat_fin) / 2)
    distancia = np.sqrt(delta_lon**2 + delta_lat**2)    
    if unidad == 'mn':
        distancia = distancia / 1.852  # Convertir kilómetros a millas náuticas
    return distancia


def calcular_tiempo_de_viaje(lon_fin, lat_fin, lon_ini, lat_ini, velocidad, unidad='km'):
    """ Entradas: 
    lon_ini = Float, longitud inicial en grados decimales
    lat_ini = Float, latitud inicial en grados decimales
    lon_fin = Float, longitud final en grados decimales
    lat_fin = Float, latitud final en grados decimales
    velocidad = Float, velocidad de viaje en km/h o nudos (dependiendo de la unidad)
    unidad = String, 'km' para kilómetros o 'mn' para millas náuticas  
    
    Salida:
    tiempo = Float, tiempo de viaje entre las coordenadas en la unidad especificada
    
    Notas: 1 km = 1.852 millas náuticas
    """
    if velocidad <= 0:
        raise ValueError("La velocidad debe ser mayor que cero.")
    
    distancia = distancia_entre_dos_puntos(lon_fin, lat_fin, lon_ini, lat_ini, unidad)
    if unidad == 'mn':
        # Convertir velocidad de nudos a km/h (1 nudo = 1.852 km/h)
        velocidad = velocidad * 1.852

    # Calcular tiempo de viaje en horas
    tiempo_horas = distancia / velocidad
    return tiempo_horas

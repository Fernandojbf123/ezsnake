import numpy as np
import datetime
import pandas as pd


# FECHAS DE MATLAB A PYTHON en formato datetime
def datenum_to_datetime(datenum: list | np.ndarray, minutes_rounded: bool = True) -> np.ndarray:
    """Lee un array en formato datenum de matlab (numeros de orden 7XXXXX) y los cambia al formato datetime de python.
    Por lo general los datos datenum vienen desde un archivo netCDF generado por matlab por lo que pueden venir como un Masked Array.
    Esta Funcion utiliza las librerías datetime y numpy para hacer la conversión de fechas.
    
    Entradas:
    datenum = array de fechas en formato datenum de matlab (puede ser un Masked Array o un array normal)
    minutes_rounded = True si se desea redondear al minuto más cercano, False si se desea mantener la precisión original (segundos y microsegundos)
    
    Salida:
    Un array de fechas en formato datetime de python, con el mismo tamaño que el array de entrada.
    ***********
    by BelloDev
    agregado 2025/04/06
    ultima revision 2026/05/25
    ***********
    """
    # Primero convierte el maskedArray en un array normal de python
    if isinstance(datenum, np.ma.MaskedArray):
        fechas = np.array(datenum.filled(np.nan))
    else:
        fechas = np.array(datenum)

    # Lista para almacenar las fechas convertidas
    fechas_convertidas = []
    
    for idatenum in fechas.flatten():
         # Conversion de idatenum a un valor escalar
        idatenum = float(idatenum)

        # Convierte un número datenum de MATLAB a datetime de Python
        fecha = datetime.datetime.fromordinal(int(idatenum)) + datetime.timedelta(days=idatenum % 1) - datetime.timedelta(days=366)
        #fechas_convertidas.append(fecha)

        # Redondear al minuto más cercano
        fecha = fecha.replace(microsecond=0)
        if fecha.second >= 30:
            fecha += datetime.timedelta(minutes=1)
        
        fecha = fecha.replace(second=0)
        fechas_convertidas.append(fecha)

    # Retorna las fechas convertidas como un array de numpy
    return np.array(fechas_convertidas)


def datenum_to_pd_datetime(datenum, minutes_rounded: bool = True):
    """ Convierte un array de datenum de MATLAB a un array de datetime de pandas.
     Entradas:
     
    datenum = array de fechas en formato datenum de matlab (puede ser un Masked Array o un array normal)
    minutes_rounded = True si se desea redondear al minuto más cercano, False si se desea mantener la precisión original (segundos y microsegundos)
    
    Salida:
    Una serie en formato pandas datetime, con el mismo tamaño que el array de entrada.
    ***********
    by BelloDev
    agregado 2026/05/25
    ultima revision 2026/05/25
    ***********
    """
    
    datetime_array = datenum_to_datetime(datenum, minutes_rounded=minutes_rounded)
    
    return pd.to_datetime(datetime_array)
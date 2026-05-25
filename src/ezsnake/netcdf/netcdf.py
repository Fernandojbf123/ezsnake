"""
Módulo para lectura, exploración y escritura de archivos NETCDF usando netCDF4.
"""

import numpy as np
from netCDF4 import Dataset
from typing import List, Union, Dict, Any, Literal

def load_nc(ruta: str, variable: Union[str, List[str], None] = None):
    """
    Carga variables de un archivo NetCDF.
    Si variable es None, retorna todo el archivo como un dict de numpy.array.
    Si es un string, retorna la variable como numpy.array.
    Si es una lista de strings, retorna una tupla de numpy.array en el mismo orden.
    """
    with Dataset(ruta, 'r') as nc:
        if variable is None:
            # Cargar todas las variables
            return {var: np.array(nc.variables[var][:]) for var in nc.variables}
        elif isinstance(variable, str):
            return np.array(nc.variables[variable][:])
        elif isinstance(variable, list):
            return tuple(np.array(nc.variables[v][:]) for v in variable)
        else:
            raise ValueError("variable debe ser None, un string o una lista de strings.")

def view_att(ruta: str, variable: str = ""):
    """
    Muestra los atributos de una variable o los atributos globales si variable es None.
    """
    with Dataset(ruta, 'r') as nc:
        if variable is None:
            return {att: getattr(nc, att) for att in nc.ncattrs()}
        else:
            if variable not in nc.variables:
                raise ValueError(f"La variable '{variable}' no existe en el archivo.")
            return {att: getattr(nc.variables[variable], att) for att in nc.variables[variable].ncattrs()}

def dict2nc(ruta_salida: str, 
            data: Dict[str, Any],
            formato: Literal['NETCDF4', 'NETCDF4_CLASSIC', 'NETCDF3_CLASSIC', 'NETCDF3_64BIT_OFFSET', 'NETCDF3_64BIT_DATA'] = 'NETCDF4') -> None:
    """
    Crea un archivo NetCDF a partir de un diccionario con atributos globales y variables.
    El diccionario debe tener la clave 'global_atributes' y una clave por cada variable.
    Cada variable debe tener las siguientes claves:
    - 'value': numpy.array con los datos de la variable
    - 'dims': lista de strings con los nombres de las dimensiones de la variable
    - 'attrs': diccionario con atributos de la variable (opcional)
        - 'units': string con las unidades de la variable (opcional)
        - 'long_name': string con el nombre largo de la variable (opcional)
        - '_FillValue': valor para representar datos faltantes (opcional)
        
    se puede usar la clase BuildNetCDFVariable para construir cada variable de forma más ordenada.
    
    dict_ejemplo = {
        'global_atributes': {
            'titulo': 'Ejemplo de archivo NetCDF',
            'institucion': 'Mi Institución',
            'fuente': 'Simulación'
        },
        'T': {
            'value': np.random.rand(10, 20),
            'dims': ['lat', 'lon'],
            'attrs':{
                'long_name': 'Temperatura',
                'units': 'K',
                '_FillValue': -9999.0
            }  
        },
        'Pa': {
            'value': np.random.rand(10, 20),
            'dims': ['lat', 'lon'],
            'attrs':{
                'long_name': 'Presion atmosferica',
                'units': 'psi'
            }
        },
        'lat': {
            'value': np.linspace(-90, 90, 10),
            'dims': ['lat'],
            'attrs':{
                'long_name': 'Latitud',
                'units': 'degrees_north'
            }
        },
        'lon': {
            'value': np.linspace(-180, 180, 20),
            'dims': ['lon'],
            'attrs':{
                'long_name': 'Longitud',
                'units': 'degrees_east'
            }
        }
    
    """
    with Dataset(ruta_salida, "w", format=formato) as ds:

        # 1. AGREGAR ATRIBUTOS GLOBALES
        global_attrs = data.get("global_atributes", {})

        for attr_name, attr_value in global_attrs.items():
            setattr(ds, attr_name, attr_value)

        # 2. DETECTAR Y CREAR DIMENSIONES
        dimensiones = {}
        for variable_name, variable_info in data.items():
            if variable_name != "global_atributes":
                values = np.asarray(variable_info["value"])
                dims = variable_info["dims"]
                for dim_name, dim_size in zip(dims, values.shape):
                    if dim_name not in dimensiones:
                        dimensiones[dim_name] = dim_size

        # Crear dimensiones en el archivo
        for dim_name, dim_size in dimensiones.items():
            ds.createDimension(dim_name, dim_size)

        # 3. CREAR VARIABLES
        for variable_name, variable_info in data.items():

            if variable_name != "global_atributes":
                values = np.asarray(variable_info["value"])
                dims = tuple(variable_info["dims"])
                attrs = variable_info.get("attrs", {})

                # Detectar tipo netCDF
                dtype = values.dtype

                # Manejo especial de _FillValue
                # Debe pasarse al crear variable
                fill_value = attrs.pop("_FillValue", None)
                if fill_value is not None:
                    var = ds.createVariable(
                        variable_name,
                        dtype,
                        dims,
                        fill_value=fill_value,
                        zlib=True)

                else:
                    var = ds.createVariable(
                        variable_name,
                        dtype,
                        dims,
                        zlib=True)

                # Escribir datos
                var[:] = values

                # Agregar atributos
                for attr_name, attr_value in attrs.items():
                    if attr_value is not None:
                        setattr(var, attr_name, attr_value)

    print(f"Archivo NetCDF guardado en: {ruta_salida}")
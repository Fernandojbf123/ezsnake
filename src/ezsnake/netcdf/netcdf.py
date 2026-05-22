"""
Módulo para lectura, exploración y escritura de archivos NETCDF usando netCDF4.
"""

import numpy as np
from netCDF4 import Dataset
from typing import List, Union, Dict, Any

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

def dict2nc(ruta: str, data: Dict[str, Any]):
    """
    Crea un archivo NetCDF a partir de un diccionario con atributos globales y variables.
    El diccionario debe tener la clave 'global_atributes' y una clave por cada variable.
    Cada variable debe tener 'value', 'dims' y atributos opcionales.
    """
    from collections import OrderedDict
    with Dataset(ruta, 'w') as nc:
        # Crear dimensiones
        dims_dict = OrderedDict()
        for var, vdata in data.items():
            if var == 'global_atributes':
                continue
            dims = vdata.get('dims', [])
            for i, d in enumerate(dims):
                if d not in dims_dict:
                    dims_dict[d] = vdata['value'].shape[i]
        for d, size in dims_dict.items():
            nc.createDimension(d, size)
        # Atributos globales
        for att, val in data.get('global_atributes', {}).items():
            setattr(nc, att, val)
        # Variables
        for var, vdata in data.items():
            if var == 'global_atributes':
                continue
            dims = vdata.get('dims', [])
            # Forzar a float todos los valores
            value = np.array(vdata['value'], dtype=float)
            # Manejar _FillValue correctamente
            fill_value = vdata.get('_FillValue', None)
            create_var_kwargs = {}
            if fill_value is not None:
                create_var_kwargs['fill_value'] = fill_value
            var_obj = nc.createVariable(var, value.dtype, dims, **create_var_kwargs)
            var_obj[:] = value
            for att, val in vdata.items():
                if att not in ['value', 'dims', '_FillValue']:
                    setattr(var_obj, att, val)

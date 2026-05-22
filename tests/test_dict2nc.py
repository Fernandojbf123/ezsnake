import os
import sys
import numpy as np

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.ezsnake.netcdf.netcdf import dict2nc, load_nc


# Datos de ejemplo
dic_de_prrueba = {
    'global_atributes': {
        'title': 'Archivo de prueba',
        'institution': 'Test Lab'
    },
    'temp': {
        'value': np.array([[35, 34], [33, 32]]),
        'dims': ['lat', 'lon'],
        'units': 'Celsius',
        '_FillValue': np.nan,
        '_missing_value': np.nan,
        'valid_range': [-10, 50]
    },
    'pres': {
        'value': np.array([1013, 1012]),
        'dims': ['lat'],
        'units': 'milibars',
        '_FillValue': np.nan,
        '_missing_value': np.nan,
        'valid_range': [500, 1250]
    }
}

def test_dict2nc_variable_exist(tmp_path):
    # Datos de ejemplo
    data = dic_de_prrueba.copy()
     
    # Ruta temporal para el archivo NetCDF
    nc_path = tmp_path / "test.nc"
    dict2nc(str(nc_path), data)

    # Verificar que el archivo se creó y contiene los datos correctos
    loaded = load_nc(str(nc_path))
    
    for nombre_de_variable in data.keys():
        if nombre_de_variable != 'global_atributes':
            assert nombre_de_variable in loaded, f"La variable '{nombre_de_variable}' no se encuentra en el archivo NetCDF."
    

def test_dict2nc_tiene_mismo_valor(tmp_path):
    data = dic_de_prrueba.copy()
    
    # Ruta temporal para el archivo NetCDF
    nc_path = tmp_path / "test.nc"

    dict2nc(str(nc_path), data)

    # Verificar que los valores cargados son iguales a los originales
    loaded = load_nc(str(nc_path))
    
    for nombre_de_variable in data.keys():
        if nombre_de_variable != 'global_atributes':
            assert np.array_equal(loaded[nombre_de_variable], data[nombre_de_variable]['value'])
    


if __name__ == "__main__":
    test_dict2nc_variable_exist("test.nc")
    print("Prueba dict2nc completada correctamente.")


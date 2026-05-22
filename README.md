# ezsnake

Librería de utilidades para procesamiento de datos científicos en Python.

## Módulos disponibles

### 📦 NetCDF (`ezsnake.netcdf`)

Módulo para lectura, exploración y escritura de archivos NetCDF usando netCDF4.

#### Funciones disponibles:

- **`load_nc(ruta, variable=None)`**
  - Carga variables de un archivo NetCDF.
  - Si `variable` es `None`, retorna todas las variables como diccionario.
  - Si `variable` es un string, retorna esa variable como numpy.array.
  - Si `variable` es una lista de strings, retorna una tupla de numpy.array.

- **`view_att(ruta, variable="")`**
  - Muestra los atributos de una variable específica o los atributos globales del archivo.
  - Si `variable` es `None`, retorna los atributos globales.

- **`dict2nc(ruta, data)`**
  - Crea un archivo NetCDF a partir de un diccionario.
  - El diccionario debe contener la clave `'global_atributes'` y una clave por cada variable.
  - Cada variable debe incluir `'value'`, `'dims'` y atributos opcionales.

### 🔧 MATLAB (`ezsnake.matlab`)

Módulo para interoperabilidad con MATLAB (en desarrollo).

### 🛠️ Utilidades (`ezsnake.utils`)

Módulo de utilidades generales (en desarrollo).

## Instalación

```bash
pip install -e .
```

## Uso básico

```python
from ezsnake.netcdf import load_nc, view_att, dict2nc

# Cargar todas las variables de un archivo NetCDF
data = load_nc('archivo.nc')

# Cargar una variable específica
temperatura = load_nc('archivo.nc', 'temperature')

# Ver atributos de una variable
attrs = view_att('archivo.nc', 'temperature')

# Crear un archivo NetCDF desde un diccionario
data_dict = {
    'global_atributes': {'title': 'Mi archivo NetCDF'},
    'temperature': {
        'value': np.array([[1, 2], [3, 4]]),
        'dims': ['time', 'lat'],
        'units': 'Celsius'
    }
}
dict2nc('nuevo_archivo.nc', data_dict)
```


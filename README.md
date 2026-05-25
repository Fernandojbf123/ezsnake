
# ezsnake

Librería de utilidades científicas en Python para NetCDF, manipulación avanzada de Word, utilidades geográficas y funciones tipo MATLAB.

---

## 📦 Módulos principales

### 1. NetCDF (`ezsnake.netcdf`)
Lectura, exploración y escritura de archivos NetCDF usando netCDF4.

**Funciones principales:**
- `load_nc(ruta, variable=None)`
- `view_att(ruta, variable=None)`
- `dict2nc(ruta, data)`

### 2. Word Template Writer (`ezsnake.word_template_writer`)
Automatización avanzada de plantillas Word (.docx): inserción de figuras, referencias cruzadas, reemplazo de variables, tablas y documentos externos.

**API pública:**
- `insertar_figuras_en_plantilla(doc, diccionario_de_reemplazos)`
- `insertar_referencias_cruzadas_en_plantilla(doc, diccionario_de_reemplazos)`
- `reemplazar_texto_en_plantilla(doc, diccionario_de_reemplazos)`
- `insertar_documento_externo_en_plantilla(doc, diccionario_de_reemplazos)`
- `rellenar_tablas_en_plantilla(doc, diccionario_de_reemplazos)`
- `reemplazar_variables_en_tablas(doc, diccionario_de_reemplazos)`

### 3. MATLAB-like (`ezsnake.matlab`)
Funciones científicas compatibles con MATLAB para estadística, distribuciones y conversión de fechas:

**Funciones principales:**
- `betacdf`, `betafit`, `betainv`, `betalike`, `betapdf`, `betaln`
- `datenum_to_datetime`, `datenum_to_pd_datetime`
- `distchck`, `hist`
- `logncdf`, `lognfit`, `lognlike`, `lognpdf`, `lognrnd`, `lognstat`
- `normcdf`, `normfit`, `normlike`, `normpdf`, `norminv`

### 4. Utilidades (`ezsnake.utils`)
Funciones para cálculos geográficos, manipulación de datos y utilidades varias:

**Funciones principales:**
- `uv2polar(u, v)`
- `polar2uv(dir_deg, spd)`
- `grados_a_km_lat(delta_lat)`
- `grados_a_km_lon(delta_lon, latitud)`
- `distancia_entre_dos_puntos(lon2, lat2, lon1, lat1, unidad='km')`
- `calcular_tiempo_de_viaje(lon2, lon1, lat2, lat1, velocidad, unidad='km')`
- `timestamp_a_texto_espanol(fecha, mes_y_anio)`
- `get_excel_variables_name(df_datos_documento)`
- `get_excel_variable_values(df_datos_documento, nombre_variable)`

---

## 🚀 Instalación

Desde PyPI:
```bash
pip install ezsnake
```

Desde el repositorio (editable):
```bash
pip install -e .
```

---

## 📋 Requerimientos

- Python >= 3.11
- netcdf4 >= 1.7.4
- numpy >= 2.0
- openpyxl >= 3.1
- pandas >= 2.0
- python-docx >= 1.1
- scipy >= 1.17.1

---

## 📝 Ejemplo de uso

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

---

## 👤 Autor y créditos

***********
by BelloDev  
agregado 2026/05/25  
ultima revision 2026/05/25
***********


# Word Template Writer

Módulo para manipular plantillas de Word (.docx) con funciones de alto nivel para insertar figuras, referencias cruzadas, texto, documentos externos y tablas.

    reemplazar_texto_en_plantilla,
    rellenar_tablas_en_plantilla,
# Aplicar transformaciones en orden

## 🚀 Uso Básico y Ejemplo de Construcción de Diccionarios

```python
import os
import pandas as pd
from docx import Document
from word_template_writer import (
    insertar_figuras_en_plantilla,
    insertar_referencias_cruzadas_en_plantilla,
    reemplazar_texto_en_plantilla,
    reemplazar_variables_en_tablas,
    insertar_documento_externo_en_plantilla,
    rellenar_tablas_en_plantilla,
)

# Abrir plantilla
doc = Document('plantilla.docx')

# Ejemplo: Construir diccionario de figuras a partir de un DataFrame
def construir_diccionario_agregar_figuras(df_datos_documento: pd.DataFrame) -> dict:
    dict_documento = {}
    varnames = get_excel_variables_name(df_datos_documento)
    for varname in varnames:
        if varname.startswith("fig_"):
            array = []
            varvalues = get_excel_variable_values(df_datos_documento, nombre_variable=varname) or []
            dict_figura = FiguraSchema(ruta_a_figura="", titulo="", tamanio=6, bookmark="")
            for fig_name in varvalues:
                ruta_a_la_carpeta_de_imagenes = ""
                dict_figura.set_ruta(ruta_a_carpeta_de_imagenes, fig_name, "jpg")
                dict_figura.set_tamanio(6)
                dict_figura.set_bookmark(fig_name)
                dict_figura.set_titulo("")
                array.append(dict_figura.to_dict())
            if array:
                dict_documento["<<"+varname+">>"] = array
    return dict_documento

# Ejemplo: Diccionario de reemplazos para texto
diccionario = {}
diccionario["<<fecha_inicio_vigencia>>"] = "2026-05-01"
diccionario["<<fecha_final_vigencia>>"] = "2026-05-31"

# Ejemplo: Diccionario de documentos externos
def construir_diccionario_de_reemplazos_para_docs_externos(diccionario_de_reemplazos: dict):
    rutas_a_carpeta = ""
    archivos = ["ejemplo1.docx", "ejemplo2.docx"]
    rutas = [os.path.join(rutas_a_carpeta, archivo) for archivo in archivos]
    diccionario_de_reemplazos["<<external_doc_plan_de_crucero>>"] = rutas

# Ejemplo: Diccionario de tablas
def construir_diccionario_de_reemplazos_para_tablas(diccionario_de_reemplazos: dict, doc: object):
    tabla1 = pd.DataFrame({})
    opciones_de_tabla = OpcionesTabla()
    estilos_de_tabla = EstilosTabla(doc)
    estilos_de_tabla.set_estilo_por_defecto("texto_tablas_centrado")
    diccionario_de_reemplazos["<<tabla_plan>>"] = {
        "tabla": tabla1,
        "estilos_de_tabla": estilos_de_tabla,
        "opciones_de_tabla": opciones_de_tabla
    }

# Aplicar transformaciones en orden
reemplazar_texto_en_plantilla(doc, diccionario)
reemplazar_variables_en_tablas(doc, diccionario)
insertar_figuras_en_plantilla(doc, diccionario)
insertar_referencias_cruzadas_en_plantilla(doc, diccionario)
insertar_documento_externo_en_plantilla(doc, diccionario)
doc.save('documento_final.docx')
```
reemplazar_variables_en_tablas(doc, diccionario)
insertar_figuras_en_plantilla(doc, diccionario)
insertar_referencias_cruzadas_en_plantilla(doc, diccionario)
insertar_documento_externo_en_plantilla(doc, diccionario)

# Guardar resultado
doc.save('documento_final.docx')
```

---

## 📋 Convenciones del Diccionario de Reemplazos

El diccionario de reemplazos debe contener palabras clave (keys) que coincidan con los marcadores en la plantilla de Word. Cada tipo de operación tiene una convención específica.

### **1. Variables de Texto** ➡️ Cualquier nombre

**Plantilla Word:**
```
Orden de servicio: <<orden_servicio>>
Cliente: <<nombre_cliente>>
Fecha: <<fecha_entrega>>
```

**Diccionario Python:**
```python
diccionario = {
    "<<orden_servicio>>": "12345",
    "<<nombre_cliente>>": "ACME Corporation",
    "<<fecha_entrega>>": "21/05/2026",
}
```

**Restricciones:**
- ❌ NO usar prefijos reservados: `fig_`, `ref_`, `tabla_`, `external_doc`
- ✅ Cualquier otro nombre es válido

---

### **2. Figuras** ➡️ Prefijo `<<fig_`

**Plantilla Word:**
```
<<fig_mapas>>
<<fig_esquema_sonda>>
<<fig_fotos>>
```

**Schema de Figuras:**

Cada figura debe seguir esta estructura:

```python
figura = {
    "ruta": "ruta/completa/a/imagen.jpg",      # Ruta absoluta o relativa a la imagen
    "titulo": "Descripción de la figura.",     # Título del pie de figura (opcional)
    "tamanio": 6,                              # Ancho en pulgadas (default: 6)
    "bookmark": "_Ref_Nombre_Figura",          # Nombre del bookmark para referencias (opcional)
}
```

**Clase Helper (opcional):**

```python
class Dictfiguras():
    def __init__(self):
        self.ruta = ""
        self.titulo = ""
        self.tamanio = 6
        self.bookmark = ""

    def set_ruta(self, varvalue: str):
        """Establece la ruta completa de la imagen"""
        carpeta = get_ruta_a_carpeta_de_las_figuras(usar_NAS=True)
        ruta_completa = os.path.join(carpeta, varvalue + ".jpg")
        self.ruta = ruta_completa.strip()
    
    def set_tamanio(self, tamanio: int):
        """Establece el ancho de la imagen en pulgadas"""
        self.tamanio = tamanio
    
    def set_bookmark(self, varvalue: str):
        """Establece el nombre del bookmark (agrega prefijo 'Ref_')"""
        bookmark = "Ref_" + varvalue
        self.bookmark = bookmark.strip()
        
    def set_titulo(self, varvalue: str):
        """Establece el título del pie de figura (agrega punto final si no existe)"""
        titulo = varvalue.strip() 
        if titulo != "":
            titulo = titulo if titulo.endswith(".") else titulo + "."
        self.titulo = titulo
        
    def return_dict(self) -> dict:
        """Retorna el diccionario con la estructura correcta"""
        return {
            "ruta": self.ruta,
            "titulo": self.titulo,
            "tamanio": self.tamanio,
            "bookmark": self.bookmark
        }
```

**Ejemplo de uso:**

```python
# Caso 1: Figuras SIN título (solo imágenes, sin Caption)
diccionario = {
    "<<fig_fotos>>": [
        {"ruta": "foto1.jpg", "titulo": "", "tamanio": 4, "bookmark": ""},
        {"ruta": "foto2.jpg", "titulo": "", "tamanio": 4, "bookmark": ""},
    ]
}

# Caso 2: Figuras CON título (con Caption, numeración automática y bookmarks)
diccionario = {
    "<<fig_mapas>>": [
        {
            "ruta": "mapa_ubicacion.jpg", 
            "titulo": "Mapa de ubicación de las sondas oceanográficas",
            "tamanio": 6, 
            "bookmark": "_Ref_Mapa_Ubicacion"
        },
        {
            "ruta": "mapa_temperatura.jpg", 
            "titulo": "Temperatura superficial del agua",
            "tamanio": 5, 
            "bookmark": "_Ref_Mapa_Temperatura"
        },
    ]
}

# Caso 3: Usando la clase helper Dictfiguras
dict_temporal = Dictfiguras()
figuras = []

for nombre_imagen in ["sonda_001", "sonda_002"]:
    dict_temporal.set_ruta(nombre_imagen)  # Se agrega .jpg automáticamente
    dict_temporal.set_tamanio(6)
    dict_temporal.set_bookmark(nombre_imagen)  # Se agrega prefijo "Ref_"
    dict_temporal.set_titulo("Esquema de la sonda oceanográfica")
    figuras.append(dict_temporal.return_dict())

diccionario = {
    "<<fig_esquema_sonda>>": figuras
}
```

**Comportamiento automático:**

- ✅ Si `titulo != ""`: Se crea pie de figura con estilo Caption
- ✅ Numeración automática con campo SEQ (Figura 1, Figura 2, ...)
- ✅ Bookmarks creados automáticamente para referencias cruzadas
- ✅ Estilo centrado si título < 115 caracteres, justificado si es mayor
- ❌ Si `titulo == ""`: Solo se insertan imágenes sin Caption ni bookmarks

**📌 Nota importante:** Después de ejecutar `insertar_figuras_en_plantilla()`, el diccionario se **muta** agregando automáticamente keys de tipo `<<ref_*>>` con los bookmarks creados.

---

### **3. Referencias Cruzadas a Figuras** ➡️ Prefijo `<<ref_`

**¿Cómo funciona?**

1. Ejecutar `insertar_figuras_en_plantilla()` primero (crea los bookmarks)
2. El diccionario se muta automáticamente agregando keys `<<ref_*>>`
3. Ejecutar `insertar_referencias_cruzadas_en_plantilla()` para insertar las referencias

**Plantilla Word:**
```
De la <<ref_mapas>> se observa la distribución espacial de las sondas.
En la <<ref_esquema_sonda>> se muestra el diseño técnico del equipo.
```

**Diccionario Python (después de insertar figuras):**

```python
# Estado INICIAL del diccionario:
diccionario = {
    "<<fig_mapas>>": [
        {"ruta": "mapa1.jpg", "titulo": "Mapa 1", "tamanio": 6, "bookmark": "_Ref_Mapa1"},
        {"ruta": "mapa2.jpg", "titulo": "Mapa 2", "tamanio": 6, "bookmark": "_Ref_Mapa2"},
    ],
    "<<fig_esquema_sonda>>": [
        {"ruta": "esquema.jpg", "titulo": "Esquema", "tamanio": 6, "bookmark": "_Ref_Esquema"},
    ]
}

# Después de ejecutar insertar_figuras_en_plantilla():
# El diccionario AHORA contiene (automáticamente):
{
    "<<fig_mapas>>": [...],  # Se mantiene
    "<<ref_mapas>>": ["_Ref_Mapa1", "_Ref_Mapa2"],  # ✨ AGREGADO AUTOMÁTICAMENTE
    "<<fig_esquema_sonda>>": [...],  # Se mantiene
    "<<ref_esquema_sonda>>": ["_Ref_Esquema"],  # ✨ AGREGADO AUTOMÁTICAMENTE
}
```

**Resultado en Word:**

```
De la Figura 1 a la 2 se observa la distribución espacial de las sondas.
                 ↑        ↑
            (clickeable) (clickeable)

En la Figura 3 se muestra el diseño técnico del equipo.
          ↑
     (clickeable)
```

**Reglas de generación:**

- 📍 **Una figura:** `<<ref_mapas>>` → "Figura X" (solo un bookmark)
- 📍 **Varias figuras:** `<<ref_mapas>>` → "Figura X a la Y" (primer y último bookmark)
- 📍 **Sin título:** `<<ref_fotos>>` → `None` (no se crean referencias)

**Nomenclatura automática:**

| Key de figura          | Key de referencia generado |
|------------------------|----------------------------|
| `<<fig_mapas>>`        | `<<ref_mapas>>`            |
| `<<fig_esquema_sonda>>`| `<<ref_esquema_sonda>>`    |
| `<<fig_pruebas_lab>>`  | `<<ref_pruebas_lab>>`      |

**💡 Tip:** El nombre del key de referencia se genera automáticamente reemplazando `fig_` por `ref_` en el key original.

---

### **4. Documentos Externos** ➡️ Prefijo `<<external_doc_`

**Plantilla Word:**
```
<<external_doc_plan>>
```

**Diccionario Python:**

```python
# Un solo documento
diccionario = {
    "<<external_doc_plan>>": "plan_crucero.docx",
}

# Múltiples documentos (se separan con salto de página)
diccionario = {
    "<<external_doc_plan>>": [
        "plan_crucero_enero.docx",
        "plan_crucero_febrero.docx",
    ],
}
```

**⚠️ Convención especial:**

Por compatibilidad histórica, la key `"<<ruta_plan_de_crucero>>"` también es aceptada:

```python
diccionario = {
    "<<ruta_plan_de_crucero>>": "plan_crucero.docx",  # ✅ También funciona
}
```

**Comportamiento:**

- ✅ Copia párrafos y tablas del documento externo
- ✅ Mantiene formato, estilos e imágenes
- ✅ Copia relaciones de imágenes correctamente
- ❌ Excluye headers, footers, marcas de agua y propiedades de documento
- 📄 Agrega salto de página entre múltiples documentos

---

### **5. Tablas** ➡️ Prefijo `<<table_`

Soporte completo para insertar DataFrames en tablas Word con estilos configurables, colores, MultiIndex y merged cells.

**Plantilla Word:**

El marcador puede estar en cualquier celda de la tabla (típicamente en la primera fila de datos):

```
┌─────────────┬──────────────┬───────┐
│ Secuencia   │ Localización │ Lat   │  ← Encabezados (pre-formateados en plantilla)
├─────────────┼──────────────┼───────┤
│ <<table_posiciones>>         │       │  ← Marcador en primera celda de datos
├─────────────┼──────────────┼───────┤
│                               ...    │
└─────────────┴──────────────┴───────┘
```

#### 🔹 **Uso Básico**

```python
import pandas as pd
from docx import Document
from word_template_writer import rellenar_tablas_en_plantilla

doc = Document('plantilla.docx')

# Opción 1: Diccionario
datos = {
    "Secuencia": [0, 1, 2],
    "Localización": ['BOT-01', 'BOT-02', 'BOT-03'],
    "Latitud": [18.5, 18.6, 18.7],
}

# Opción 2: DataFrame
df = pd.DataFrame(datos)

# Insertar con estilos por defecto
rellenar_tablas_en_plantilla(doc, "<<table_posiciones>>", datos)
doc.save('resultado.docx')
```

**Comportamiento por defecto:**
- ✅ Aplica estilo "Normal" a todas las celdas
- ✅ Alineación vertical centrada
- ✅ Altura de fila ~4 cm
- ✅ Aplana MultiIndex automáticamente
- ✅ Elimina fila marcador después de insertar datos
- ❌ No combina celdas

---

#### 🎨 **Uso Avanzado con Estilos Personalizados**

```python
from docx import Document
from word_template_writer import (
    rellenar_tablas_en_plantilla,
    EstilosTabla,
    OpcionesTabla,
    get_estilos_disponibles
)
import pandas as pd

doc = Document('plantilla.docx')

# Ver estilos disponibles en la plantilla
estilos_disponibles = get_estilos_disponibles(doc)
print(estilos_disponibles)  # ['Normal', 'texto_tablas_centrado', 'texto_tablas_justificado', ...]

# ===== CONFIGURAR ESTILOS =====
estilos = EstilosTabla(doc)

# Estilos por defecto (se aplican a todas las celdas)
estilos.set_estilo_por_defecto('texto_tablas_centrado')
estilos.set_color_fondo_por_defecto(None)  # Sin color
estilos.set_alineacion_vertical_por_defecto('center')
estilos.set_altura_fila_por_defecto(288290)  # ~4cm en twips

# Estilos por columna (sobrescribe defecto)
estilos.set_color_de_columna(0, (230, 230, 250))  # Lavanda columna 0
estilos.set_estilo_de_columna(1, 'texto_tablas_justificado')  # Justificado columna 1
estilos.set_color_de_texto_de_columna(2, (255, 0, 0))  # Texto rojo columna 2
estilos.set_alineacion_horizontal_de_columna(3, 'right')  # Alinear derecha columna 3

# Estilos por fila (sobrescribe columna y defecto)
estilos.set_color_de_fila(0, (255, 240, 245))  # Rosado primera fila
estilos.set_estilo_de_fila(2, 'texto_tablas_justificado')  # Fila 2 justificada

# Estilos por celda específica (sobrescribe todo)
estilos.set_color_de_celda((1, 2), (173, 216, 230))  # Azul claro fila 1, col 2
estilos.set_estilo_de_celda((0, 0), 'texto_tablas_centrado')  # Celda (0,0)

# ===== CONFIGURAR OPCIONES =====
opciones = OpcionesTabla()

opciones.set_aplanar_multiindex(True)  # Aplica reset_index() a MultiIndex automáticamente
opciones.set_detectar_merge(True)  # Combina celdas con valores repetidos consecutivos
opciones.set_columnas_para_merge([0, 1])  # Solo merge en columnas 0 y 1 (None = todas)
opciones.set_eliminar_fila_marcador(True)  # Elimina fila con marcador

# ===== DATOS =====
df = pd.DataFrame({
    'Tipo': ['A', 'A', 'A', 'B', 'B'],
    'Subtipo': ['A1', 'A1', 'A2', 'B1', 'B1'],
    'Valor': [10, 20, 30, 40, 50],
    'Comentario': ['Bajo', 'Medio', 'Alto', 'Bajo', 'Medio']
})

# ===== INSERTAR TABLA =====
rellenar_tablas_en_plantilla(doc, "<<table_datos>>", df, estilos, opciones)

doc.save('resultado.docx')
```

**Resultado visual:**

```
┌─────────┬─────────┬───────┬──────────────┐
│   Tipo  │ Subtipo │ Valor │  Comentario  │  ← Encabezados (pre-formateados)
├─────────┼─────────┼───────┼──────────────┤
│    A    │   A1    │   10  │    Bajo      │  ← Fila 0 rosada
│    ║    │   ║     │   20  │    Medio     │  ← Merge automático en cols 0-1
│    ║    │   A2    │   30  │    Alto      │
│    B    │   B1    │   40  │    Bajo      │
│    ║    │   ║     │   50  │    Medio     │
└─────────┴─────────┴───────┴──────────────┘
     ↑         ↑
   Lavanda  Justificado
```

---

#### 📚 **Referencia Rápida: EstilosTabla**

```python
estilos = EstilosTabla(doc)  # Requiere documento para validar estilos

# ===== Por Defecto =====
estilos.set_estilo_por_defecto('Normal')
estilos.set_color_fondo_por_defecto((255, 255, 240))  # Amarillo claro
estilos.set_alineacion_vertical_por_defecto('center')  # 'center', 'top', 'bottom'
estilos.set_altura_fila_por_defecto(288290)  # Twips (~4cm)

# ===== Por Columna =====
estilos.set_estilo_de_columna(0, 'texto_tablas_centrado')
estilos.set_color_de_columna(0, (230, 230, 250))  # RGB 0-255
estilos.set_color_de_texto_de_columna(0, (0, 0, 0))  # Negro
estilos.set_alineacion_horizontal_de_columna(0, 'center')  # 'left', 'center', 'right', 'justify'

# ===== Por Fila =====
estilos.set_color_de_fila(0, (255, 240, 245))
estilos.set_estilo_de_fila(0, 'texto_tablas_centrado')

# ===== Por Celda =====
estilos.set_color_de_celda((1, 2), (173, 216, 230))  # (fila, columna)
estilos.set_estilo_de_celda((1, 2), 'texto_tablas_justificado')

# ===== Getters =====
config_dict = estilos.to_dict()
estilo_col = estilos.get_estilo_por_defecto()
color_col = estilos.get_color_de_columna(0)

# ===== Reset =====
estilos.reset()  # Limpia toda configuración
estilos.set_default()  # Restaura valores por defecto razonables
```

**Jerarquía de aplicación:** Celda > Fila > Columna > Defecto

---

#### 📚 **Referencia Rápida: OpcionesTabla**

```python
opciones = OpcionesTabla()

# ===== MultiIndex =====
opciones.set_aplanar_multiindex(True)  # reset_index() automático si True
# Si False y hay MultiIndex → Error

# ===== Merged Cells =====
opciones.set_detectar_merge(True)  # Combina celdas con valores repetidos
opciones.set_columnas_para_merge([0, 1])  # None = todas, [] = ninguna, [0,1] = solo 0 y 1

# ===== Fila Marcador =====
opciones.set_eliminar_fila_marcador(True)  # Elimina fila del marcador después de insertar

# ===== Getters =====
opciones_dict = opciones.to_dict()
aplanar = opciones.get_aplanar_multiindex()
detectar = opciones.get_detectar_merge()

# ===== Reset =====
opciones.reset()  # Limpia opciones
opciones.set_default()  # Restaura por defecto
```

---

#### 🔍 **Compatibilidad con Diccionarios (Backward Compatibility)**

Si prefieres no usar clases, puedes pasar diccionarios directamente:

```python
config_dict = {
    "por_defecto": {
        "estilo_parrafo": "Normal",
        "alineacion_vertical": "center",
        "altura_fila": 288290,
        "color_fondo": None
    },
    "por_columna": {
        0: {"color_fondo": (230, 230, 250)},
        1: {"estilo_parrafo": "texto_tablas_justificado"}
    },
    "por_fila": {
        0: {"color_fondo": (255, 240, 245)}
    },
    "por_celda": {
        (1, 2): {"color_fondo": (173, 216, 230)}
    }
}

opciones_dict = {
    "aplanar_multiindex": True,
    "detectar_merge": True,
    "columnas_para_merge": [0, 1],
    "eliminar_fila_marcador": True
}

rellenar_tablas_en_plantilla(doc, "<<table_datos>>", df, config_dict, opciones_dict)
```

---

#### ⚙️ **Notas Técnicas**

**Colores:**
- Formato RGB: `(R, G, B)` con valores 0-255
- Los colores se aplican vía XML (no hay API directa en python-docx)
- Ejemplo: `(230, 230, 250)` = Lavanda

**Estilos:**
- Los estilos deben existir en la plantilla Word
- Se validan al crear `EstilosTabla(doc)`
- Usa `get_estilos_disponibles(doc)` para ver estilos disponibles

**Mapeo de Columnas:**
- Se hace por ORDEN, no por nombre
- Columna 0 del DataFrame → Columna 0 de la tabla
- Asegúrate de que el DataFrame tenga el mismo número de columnas que la tabla

**Marcador:**
- Puede estar en cualquier celda de la tabla
- Típicamente se coloca en la primera celda de la primera fila de datos
- Los encabezados deben estar pre-formateados en la plantilla

**Merged Cells:**
- Solo funciona verticalmente (no horizontalmente)
- Combina celdas con valores consecutivos idénticos
- Ejemplo: ['A', 'A', 'B'] → A ocupa 2 filas, B ocupa 1

---

#### 🔄 **Tablas Semi-estáticas (Reemplazo de Variables)**

Para tablas donde **NO** necesitas llenar con un DataFrame completo, sino solo **reemplazar variables individuales** en celdas específicas:

**Función:** `reemplazar_variables_en_tablas(doc, diccionario_de_reemplazos)`

**Caso de uso típico:**
- Columna 0: Textos fijos ("Serial", "Profundidad máxima", "Precisión")
- Columna 1: Variables individuales (`<<serial>>`, `<<profundidad>>`, `<<precision>>`)
- Encabezado: Variable (`<<nombre_equipo>>`)

**Ejemplo:**

```python
from docx import Document
from word_template_writer import reemplazar_variables_en_tablas

doc = Document('plantilla.docx')

diccionario = {
    "<<nombre_equipo>>": "Sonda CTD #1",
    "<<serial>>": "4878505",
    "<<profundidad>>": "200 m",
    "<<precision>>": "±0.01°C",
    "<<fecha_calibracion>>": "15/05/2026"
}

# Busca y reemplaza variables en TODAS las tablas del documento
reemplazar_variables_en_tablas(doc, diccionario)

doc.save('resultado.docx')
```

**Plantilla Word (antes):**

```
┌──────────────────────────────┬─────────────────────┐
│        <<nombre_equipo>>                           │  ← Encabezado mergeado
├──────────────────────────────┼─────────────────────┤
│ Serial                       │ <<serial>>          │
│ Profundidad máxima           │ <<profundidad>>     │
│ Precisión                    │ <<precision>>       │
│ Fecha calibración            │ <<fecha_calibracion>>│
└──────────────────────────────┴─────────────────────┘
```

**Resultado Word (después):**

```
┌──────────────────────────────┬─────────────────────┐
│        Sonda CTD #1                                │
├──────────────────────────────┼─────────────────────┤
│ Serial                       │ 4878505             │
│ Profundidad máxima           │ 200 m               │
│ Precisión                    │ ±0.01°C             │
│ Fecha calibración            │ 15/05/2026          │
└──────────────────────────────┴─────────────────────┘
```

**Comportamiento:**
- ✅ Busca en **todas las tablas** del documento automáticamente
- ✅ Preserva formato original del texto
- ✅ Ignora variables con prefijos: `<<fig_`, `<<ref_`, `<<tabla_`, `<<external_doc`
- ✅ Reutiliza la misma lógica de `reemplazar_texto_en_plantilla()`

**¿Cuándo usar cada función?**

| Función | Uso |
|---------|-----|
| `reemplazar_texto_en_plantilla()` | Variables en **párrafos** del documento |
| `reemplazar_variables_en_tablas()` | Variables en **celdas de tablas** (semi-estáticas) |
| `rellenar_tablas_en_plantilla()` | Llenar tabla completa con **DataFrame** (dinámicas) |

---

## 🔄 Orden de Ejecución Recomendado

```python
from docx import Document
from word_template_writer import *

doc = Document('plantilla.docx')
diccionario = {...}

# 1. PRIMERO: Reemplazar texto en párrafos (más rápido, sin mutaciones)
reemplazar_texto_en_plantilla(doc, diccionario)

# 2. SEGUNDO: Reemplazar variables en tablas semi-estáticas
reemplazar_variables_en_tablas(doc, diccionario)

# 3. TERCERO: Insertar figuras (muta el diccionario agregando <<ref_*>>)
insertar_figuras_en_plantilla(doc, diccionario)

# 4. CUARTO: Insertar referencias cruzadas (usa los bookmarks del paso 3)
insertar_referencias_cruzadas_en_plantilla(doc, diccionario)

# 5. QUINTO: Insertar documentos externos (puede alterar numeración)
if "<<external_doc_plan>>" in diccionario:
    insertar_documento_externo_en_plantilla(doc, diccionario)

# 6. SEXTO: Rellenar tablas dinámicas con DataFrames (si aplica)
if "<<tabla_datos>>" in diccionario:
    rellenar_tablas_en_plantilla(doc, diccionario)

doc.save('resultado.docx')
```

**⚠️ Importante:** Seguir este orden evita conflictos y asegura que las referencias cruzadas funcionen correctamente.

---

## 📝 Ejemplo Completo

```python
import os
from docx import Document
from word_template_writer import (
    insertar_figuras_en_plantilla,
    insertar_referencias_cruzadas_en_plantilla,
    reemplazar_texto_en_plantilla,
    insertar_documento_externo_en_plantilla,
)

# Crear diccionario de reemplazos
diccionario = {
    # Texto simple
    "<<orden_servicio>>": "OS-2026-001",
    "<<cliente>>": "Instituto Oceanográfico",
    "<<fecha_entrega>>": "21/05/2026",
    "<<numero_de_sondas>>": 3,
    
    # Figuras CON título (se numeran automáticamente)
    "<<fig_mapas>>": [
        {
            "ruta": "mapa_ubicacion.jpg",
            "titulo": "Ubicación de las sondas oceanográficas en el Caribe",
            "tamanio": 6,
            "bookmark": "_Ref_Mapa_Ubicacion"
        },
        {
            "ruta": "mapa_temperatura.jpg",
            "titulo": "Distribución de temperatura superficial del mar",
            "tamanio": 6,
            "bookmark": "_Ref_Mapa_Temperatura"
        },
    ],
    
    # Figuras SIN título (solo imágenes)
    "<<fig_fotos>>": [
        {"ruta": "foto1.jpg", "titulo": "", "tamanio": 4, "bookmark": ""},
        {"ruta": "foto2.jpg", "titulo": "", "tamanio": 4, "bookmark": ""},
    ],
    
    # Documento externo
    "<<ruta_plan_de_crucero>>": "plan_crucero_mayo_2026.docx",
}

# Abrir plantilla
doc = Document('plantilla_reporte.docx')

# Aplicar transformaciones
reemplazar_texto_en_plantilla(doc, diccionario)
insertar_figuras_en_plantilla(doc, diccionario)
insertar_referencias_cruzadas_en_plantilla(doc, diccionario)
insertar_documento_externo_en_plantilla(doc, diccionario)

# Guardar resultado
doc.save('reporte_final.docx')

print("✅ Documento generado exitosamente")
print(f"📊 Referencias creadas: {diccionario.get('<<ref_mapas>>', [])}")
```

---

## 🛠️ Utilidades Adicionales

### `insert_line_feed()`

Inserta un nuevo párrafo después del párrafo dado usando manipulación XML.

```python
from word_template_writer import insert_line_feed

# Agregar párrafo vacío
nuevo_p = insert_line_feed(parrafo_actual)

# Agregar párrafo con texto
nuevo_p = insert_line_feed(parrafo_actual, texto="Texto del nuevo párrafo")

# Agregar párrafo centrado
nuevo_p = insert_line_feed(parrafo_actual, texto="Título", centrado=True)
```

---

## 📚 Estructura del Módulo

```
word_template_writer/
├── __init__.py                 # API pública
├── api.py                      # 5 funciones orquestadoras (español)
├── _figure_helpers.py          # Helpers para figuras (privado)
├── _text_helpers.py            # Helpers para texto (privado)
├── _document_helpers.py        # Helpers para documentos (privado)
├── _table_helpers.py           # Helpers para tablas (privado)
├── utils.py                    # Utilidades públicas
└── README.md                   # Este archivo
```

**Convención:** Archivos con prefijo `_` son privados y no deben importarse directamente.

---

## ⚠️ Palabras Clave Reservadas

**NO usar estos prefijos en variables de texto:**

- `<<fig_*>>` - Reservado para figuras
- `<<ref_*>>` - Reservado para referencias cruzadas (generado automáticamente)
- `<<external_doc_*>>` - Reservado para documentos externos
- `<<ruta_plan_de_crucero>>` - Reservado para documentos externos (legacy)
- `<<tabla*>>` - Reservado para tablas

---

## 🔧 Solución de Problemas

### Las referencias cruzadas muestran "XX" en lugar del número

**Causa:** Word no ha actualizado los campos automáticamente.

**Solución:** Abrir el documento en Word y presionar `Ctrl + A` (seleccionar todo) + `F9` (actualizar campos).

### Las imágenes no se insertan

**Causa:** La ruta del archivo es incorrecta o el archivo no existe.

**Solución:** Verificar que `os.path.exists(ruta)` retorna `True`.

### Los bookmarks no se crean

**Causa:** El `titulo` está vacío (`titulo == ""`).

**Solución:** Asegurarse de que `titulo` tenga contenido si se desean bookmarks.

### El documento externo no se inserta

**Causa:** La key `"<<ruta_plan_de_crucero>>"` no existe en el diccionario.

**Solución:** Verificar que la key esté presente antes de llamar la función:

```python
if "<<ruta_plan_de_crucero>>" in diccionario:
    insertar_documento_externo_en_plantilla(doc, diccionario)
```

---

## 📄 Licencia

Este módulo es parte del proyecto Drift Buoys. Para uso interno.

---

## 👥 Contribuciones

Para reportar problemas o sugerir mejoras, contactar al equipo de desarrollo.

**Versión:** 1.0.0  
**Última actualización:** Mayo 2026

# Word Template Writer

Modulo para manipular plantillas de Word (.docx) con funciones de alto nivel para:
- reemplazo de texto
- insercion de figuras
- referencias cruzadas de figuras y tablas
- insercion de documentos externos
- edicion/relleno de tablas

## API publica

```python
from word_template_writer import (
    reemplazar_texto_en_plantilla,
    reemplazar_variable_por_figura,
    reemplazar_referencias_cruzadas_de_figuras,
    reemplazar_variable_por_tabla,
    reemplazar_referencias_cruzadas_de_tablas,
    reemplazar_variables_en_tablas,
    insertar_documento_externo_en_plantilla,
    insertar_lista_en_plantilla,
    rellenar_tablas_en_plantilla,
)
```

## Convenciones de keys

- Texto libre: cualquier key no reservada
- Figuras: `<<fig...>>`
- Referencias de figuras: `<<reffigura...>>`
- Tablas nuevas (con titulo y bookmark): `<<nuevatabla_...>>`
- Tablas a editar/rellenar: `<<editartabla...>>`
- Referencias de tablas: `<<reftabla_...>>`
- Documentos externos: `<<external_doc_...>>`
- Listas: `<<lista_...>>`

## Flujo recomendado

```python
from docx import Document

doc = Document("plantilla.docx")
diccionario = {...}

reemplazar_texto_en_plantilla(doc, diccionario)
reemplazar_variables_en_tablas(doc, diccionario)
reemplazar_variable_por_figura(doc, diccionario)
reemplazar_referencias_cruzadas_de_figuras(doc, diccionario)
reemplazar_variable_por_tabla(doc, diccionario)
reemplazar_referencias_cruzadas_de_tablas(doc, diccionario)
insertar_documento_externo_en_plantilla(doc, diccionario)
rellenar_tablas_en_plantilla(doc, diccionario)

doc.save("resultado.docx")
```

## Ejemplos por funcion

### 1) reemplazar_texto_en_plantilla

Reemplaza solo variables de texto. Esta funcion ignora keys reservadas para figuras, tablas, referencias y documentos externos.

```python
diccionario = {
    "<<orden_servicio>>": "OS-2026-001",
    "<<cliente>>": "ACME",
    "<<fig_mapa>>": [...],              # ignorada por esta funcion
    "<<nuevatabla_resumen>>": {...},    # ignorada por esta funcion
    "<<reftabla_resumen>>": ["Reftabla_Resumen_1"],  # ignorada
}
reemplazar_texto_en_plantilla(doc, diccionario)
```

Caso especial: listas de tuplas para crear multiples parrafos con estilo.

```python
diccionario = {
    "<<seccion_1>>": [
        ("este es el primer párrafo", "estilo 2"),
        ("este es el segundo párrafo", ""),
        ("", ""),
        ("tercer párrafo", "estilo 3"),
    ]
}

# Reglas:
# - (texto, "NombreEstilo") aplica ese estilo al párrafo
# - (texto, "") usa estilo Normal
# - ("", "") inserta párrafo vacío con estilo Normal
reemplazar_texto_en_plantilla(doc, diccionario)
```

### 2) reemplazar_variable_por_figura

```python
diccionario = {
    "<<fig_mapas>>": [
        {
            "ruta": "mapa1.png",
            "titulo": "Mapa de ubicacion",
            "tamanio": 6,
            "bookmark": "RefFigura_Mapa1",
            "estilo_figura": "Figura",
            "estilo_titulo": "Normal",
        }
    ]
}

reemplazar_variable_por_figura(doc, diccionario)
# agrega automaticamente <<reffigura_mapas>> en el diccionario
```

### 3) reemplazar_referencias_cruzadas_de_figuras

```python
diccionario = {
    "<<reffigura_mapas>>": ["RefFigura_Mapa1", "RefFigura_Mapa2"],
}
reemplazar_referencias_cruzadas_de_figuras(doc, diccionario)
```

### 4) reemplazar_variable_por_tabla

Inserta titulo con bookmark y rellena la tabla objetivo.

```python
import pandas as pd
from word_template_writer import EstilosTabla

df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
estilos = EstilosTabla(doc).to_dict()

diccionario = {
    "<<nuevatabla_resumen>>": {
        "tabla": df,
        "estilos_de_tabla": estilos,
        "titulo": "Tabla 1. Resumen",
        "bookmark": "Reftabla_Resumen_1",
    }
}

reemplazar_variable_por_tabla(doc, diccionario)
# agrega automaticamente <<reftabla_resumen>> en el diccionario
```

### 5) reemplazar_referencias_cruzadas_de_tablas

```python
diccionario = {
    "<<reftabla_resumen>>": ["Reftabla_Resumen_1"],
}
reemplazar_referencias_cruzadas_de_tablas(doc, diccionario)
```

### 6) rellenar_tablas_en_plantilla

Rellena tablas cuyas keys empiecen con `<<editartabla`.

```python
import pandas as pd
from word_template_writer import EstilosTabla, OpcionesTabla

df = pd.DataFrame({"Secuencia": [0, 1], "Localizacion": ["BOT-01", "BOT-02"]})
estilos = EstilosTabla(doc)
opciones = OpcionesTabla()

diccionario = {
    "<<editartabla_posiciones>>": {
        "tabla": df,
        "estilos_de_tabla": estilos,
        "opciones_de_tabla": opciones,
    }
}

rellenar_tablas_en_plantilla(doc, diccionario)
```

### 7) reemplazar_variables_en_tablas

Para tablas semi-estaticas con placeholders en celdas.

```python
diccionario = {
    "<<serial>>": "4878505",
    "<<profundidad>>": "200 m",
}
reemplazar_variables_en_tablas(doc, diccionario)
```

### 8) insertar_documento_externo_en_plantilla

```python
diccionario = {
    "<<external_doc_plan>>": ["plan1.docx", "plan2.docx"],
}
insertar_documento_externo_en_plantilla(doc, diccionario)
```

### 9) insertar_lista_en_plantilla

```python
diccionario = {
    "<<lista_hallazgos>>": ["Hallazgo A", "Hallazgo B", "Hallazgo C"],
}
insertar_lista_en_plantilla(doc, diccionario)
```

## Notas

- Si usas referencias cruzadas en Word, actualiza campos al abrir el documento (por ejemplo, seleccionar todo y actualizar campos).
- Para tablas con estilos personalizados, el nombre del estilo debe existir en la plantilla Word.

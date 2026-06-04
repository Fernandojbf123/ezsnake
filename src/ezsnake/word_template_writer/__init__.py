"""
Word Template Writer Module
===========================

Módulo para manipular plantillas de Word (.docx) con funciones de alto nivel
para insertar figuras, referencias cruzadas, texto, documentos externos y tablas.

API Pública (en español):
    - reemplazar_variable_por_figura: Inserta imágenes con o sin títulos/captions
    - reemplazar_referencias_cruzadas_de_figuras: Crea referencias cruzadas a figuras
    - reemplazar_variable_por_tabla: Inserta tablas dinámicas y prepara referencias cruzadas
    - reemplazar_referencias_cruzadas_de_tablas: Crea referencias cruzadas a tablas
    - reemplazar_texto_en_plantilla: Reemplaza variables de texto en la plantilla
    - insertar_lista_en_plantilla: Inserta listas con viñetas en la plantilla
    - insertar_documento_externo_en_plantilla: Inserta documentos Word externos
    - rellenar_tablas_en_plantilla: Rellena tablas con datos de DataFrames
    
Clases de Configuración:
    - EstilosTabla: Configuración de estilos para tablas (colores, alineación, estilos)
    - OpcionesTabla: Opciones de comportamiento para tablas (merge, MultiIndex)
    - FigSchema: Configuración individual de figuras
    - TablaSchema: Configuración individual de tablas dinámicas
    
Utilidades:
    - insert_line_feed: Inserta nuevos párrafos usando XML
    - get_estilos_disponibles: Extrae estilos de párrafo disponibles en un documento

Uso típico básico:
    from word_template_writer import reemplazar_variable_por_figura, reemplazar_texto_en_plantilla
    from docx import Document
    
    doc = Document('plantilla.docx')
    diccionario = {
        "<<orden_servicio>>": "12345",
        "<<fig_mapas>>": [
            {"ruta": "mapa1.png", "titulo": "Mapa", "tamanio": 6, "bookmark": "RefFigura_Mapa1"}
        ]
    }
    
    reemplazar_texto_en_plantilla(doc, diccionario)
    reemplazar_variable_por_figura(doc, diccionario)
    doc.save('resultado.docx')

Uso avanzado con estilos de tabla:
    from word_template_writer import rellenar_tablas_en_plantilla, EstilosTabla, OpcionesTabla
    from docx import Document
    import pandas as pd
    
    doc = Document('plantilla.docx')
    
    # Configurar estilos
    estilos = EstilosTabla(doc)
    estilos.set_color_de_columna(0, (230, 230, 250))  # Lavanda
    estilos.set_estilo_de_columna(1, 'texto_tablas_justificado')
    
    # Configurar opciones
    opciones = OpcionesTabla()
    opciones.set_detectar_merge(True)
    opciones.set_columnas_para_merge([0, 1])
    
    # Datos
    df = pd.DataFrame({
        'Tipo': ['A', 'A', 'B'],
        'Valor': [10, 20, 30]
    })
    
    # Rellenar tabla
    diccionario = {
        "<<editartabla_datos>>": {
            "tabla": df,
            "estilos_de_tabla": estilos,
            "opciones_de_tabla": opciones,
        }
    }
    rellenar_tablas_en_plantilla(doc, diccionario)
    doc.save('resultado.docx')

Dependencias:
    - python-docx
    - pandas (solo para rellenar_tablas_en_plantilla)
"""

# Funciones principales del orquestador (API pública en español)
from .api import (
    reemplazar_variable_por_figura,
    reemplazar_referencias_cruzadas_de_figuras,
    reemplazar_variable_por_tabla,
    reemplazar_referencias_cruzadas_de_tablas,
    reemplazar_texto_en_plantilla,
    insertar_lista_en_plantilla,
    reemplazar_variable_en_tabla,
    reemplazar_variables_en_tablas,
    insertar_documento_externo_en_plantilla,
    rellenar_tablas_en_plantilla,
)

# Clases de configuración
from .schemas_helpers import (
    EstilosTabla,
    OpcionesTabla,
    FigSchema,
    TablaSchema,
    get_estilos_disponibles,
)

# Utilidades genéricas públicas
from .utils import insert_line_feed

__all__ = [
    # API principal (9 funciones orquestadoras)
    'reemplazar_variable_por_figura',
    'reemplazar_referencias_cruzadas_de_figuras',
    'reemplazar_variable_por_tabla',
    'reemplazar_referencias_cruzadas_de_tablas',
    'reemplazar_texto_en_plantilla',
    'insertar_lista_en_plantilla',
    'reemplazar_variable_en_tabla',
    'reemplazar_variables_en_tablas',
    'insertar_documento_externo_en_plantilla',
    'rellenar_tablas_en_plantilla',
    # Clases de configuración
    'EstilosTabla',
    'OpcionesTabla',
    'FigSchema',
    'TablaSchema',
    # Utilidades
    'insert_line_feed',
    'get_estilos_disponibles',
]

__version__ = '2.0.0'
__author__ = 'BelloDev'

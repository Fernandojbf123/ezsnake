import os
import pandas as pd
import numpy as np

from ..utils.utils import get_excel_variables_name, get_excel_variable_values
from .schemas_helpers import FigSchema, OpcionesTabla, EstilosTabla

############################ DICCIONARIO DE REEMPLAZOS PARA FIGURAS ############################

# ESQUEMA DEL DICCIONARIO DE FIGURAS

    

def construir_diccionario_agregar_figuras(df_datos_documento: pd.DataFrame) -> dict:
    """
    Construye un diccionario con las figuras a agregar al documento, a partir de un dataframe de pandas con los datos del documento.
    El diccionario de cada figura tiene la estructura que se indica en el esquema FigSchema.
    Si la variable tiene una o más figuras éstas se agrupan en un array de diccionarios; un diccionario por cada figura.
    Recuerda que la variable debe comenzar por "fig_" para ser reconocida como una variable que contiene las figuras.
    """
    dict_documento = {}
    varnames = get_excel_variables_name(df_datos_documento) # Acá poner la función que lee el excel de datos del documento.

    for ivarname, varname in enumerate(varnames):
        if varname.startswith("fig_"):
            array=[] # Acá se guardarán los diccionarios de cada figura asociada a la variable.
            varvalues = get_excel_variable_values(df_datos_documento = df_datos_documento, nombre_variable = varname) or []
            dict_figura = FigSchema(ruta_a_figura = "", titulo = "", tamanio = 6, bookmark = "") # uso el esquema (es una guia para construir el diccionario de cada figura)
            
            for fig_name in varvalues:
                ruta_a_la_carpeta_de_imagenes = ""
                dict_figura.set_ruta(ruta_a_la_carpeta_de_imagenes = ruta_a_la_carpeta_de_imagenes, nombre_de_archivo = fig_name, extension = "jpg")
                dict_figura.set_tamanio(6) 
                dict_figura.set_bookmark(fig_name)
                dict_figura.set_titulo("") # se puede leer de la siguiente variable del excel o colocarse manualmente acá.
                
                array.append(dict_figura.to_dict())
        
            if array != []:        
                dict_documento["<<"+varname+">>"] = array
            
    return dict_documento


############################ DICCIONARIO DE REEMPLAZOS PARA TEXTO ############################
def construir_diccionario_de_datos_documento(diccionario_de_reemplazos: dict):

    ## Ejemplos
    diccionario_de_reemplazos["<<fecha_inicio_vigencia>>"] = "ejemplo1"
    diccionario_de_reemplazos["<<fecha_final_vigencia>>"] = "ejemplo2"
    
    
############################ DICCIONARIO DE REEMPLAZOS PARA DOCS EXTERNOS ############################
def construir_diccionario_de_reemplazos_para_docs_externos(diccionario_de_reemplazos: dict):

    rutas_a_carpeta = ""
    archivos = ["ejemplo1.docx", "ejemplo2.docx", "ejemplo3.docx"] # Acá se pueden leer los nombres de los archivos desde el excel de datos del documento, por ejemplo, con la función get_excel_variable_values.
    rutas = []  
    for archivo in archivos:
        ruta = os.path.join(rutas_a_carpeta, archivo)
        rutas.append(ruta)
    
    diccionario_de_reemplazos["<<external_doc_plan_de_crucero>>"] = rutas

    
############################# DICCIONARIO DE REEMPLAZOS PARA TABLAS ############################
# Es probable que acá necesite varios esquemas, dependiendo de la tabla.
def construir_diccionario_de_reemplazos_para_tablas(diccionario_de_reemplazos: dict,
                                                    doc: object):
    
    
    # Ejemplo de una tabla que debe estar combinada por filas en las columnas 0 y 1.
    #Esta tabla sería algo como:
    #secuencia | serial_general         | equipo        | numero_de_serie_de_equipo
    #1         | serial_general_1       | texto_fijo1   | serie_equipo1
    #1         | serial_general_1       | texto_fijo2   | serie_equipo2
    #2         | serial_general_2       | texto_fijo1   | serie_equipo1
    #2         | serial_general_2       | texto_fijo2   | serie_equipo2
    #3         | serial_general_3       | texto_fijo1   | serie_equipo1
    #3         | serial_general_3       | texto_fijo2   | serie_equipo2
    
    # Ejemplo 1: Mostrar la tabla tal cual se carga
    tabla1 = pd.DataFrame({}) #supongamos que se cargo la tabla del ejemplo 
    opciones_de_tabla = OpcionesTabla() # Llamar las opciones de la tabla (acá se configura por defecto)
    estilos_de_tabla = EstilosTabla(doc) # Llamar los estilos de la tabla (acá se configura por defecto)
    estilos_de_tabla.set_estilo_por_defecto("texto_tablas_centrado") # elijo un estilo (tiene que estar denifido en la plantilla de word)
    tabla1 = pd.DataFrame({}) # Un dataframe
    
    diccionario_de_reemplazos["<<tabla_plan>>"] = {
        "tabla": tabla1,
        "estilos_de_tabla": estilos_de_tabla,
        "opciones_de_tabla": opciones_de_tabla
    }
    

    # Ejemplo2: Mostrar la tabla con filas combinadas
    opciones_de_tabla = OpcionesTabla() # Llamar las opciones de la tabla (acá se configura por defecto)
    estilos_de_tabla = EstilosTabla(doc) # Llamar los estilos de la tabla (acá se configura por defecto)
    # modifico opciones de la tabla para hacer el merge de las las filas en las columnas que corresponda 
    opciones_de_tabla.set_detectar_merge(True) # Le pido que detecte la combinacion de filas
    opciones_de_tabla.set_columnas_para_merge([0,1]) # Le indico las columnas donde se debe hacer el merge 
    estilos_de_tabla.set_estilo_de_columna(2, "texto_tablas_justificado") # indico estilos para otra columna
    estilos_de_tabla.set_estilo_de_columna(3, "texto_tablas_justificado") # indico estilos para otra columna
    
    #la tabla que se escribirá en el word sera algo así (cambiará la letra, el centrado, etc, dependiendo del estilo asignado a cada columna):
    #secuencia | serial_general         | equipo        | numero_de_serie_de_equipo
    #1         | serial_general_1       | texto_fijo1   | serie_equipo1
    #          |                        | texto_fijo2   | serie_equipo2
    #2         | serial_general_2       | texto_fijo1   | serie_equipo1
    #          |                        | texto_fijo2   | serie_equipo2
    #3         | serial_general_3       | texto_fijo1   | serie_equipo1
    #          |                        | texto_fijo2   | serie_equipo2
    
    # o sea, se agruparon los valores
    
    diccionario_de_reemplazos["<<tabla_equipos>>"] = {
        "tabla": tabla1,
        "estilos_de_tabla": estilos_de_tabla,
        "opciones_de_tabla": opciones_de_tabla
    }

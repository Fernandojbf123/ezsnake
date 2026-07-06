import ezsnake.word_template_writer as ezw
from docx import Document
import pandas as pd


ruta_plantilla =  "C:\\programacion\\codigos_python\\ezsnake\\tests\\demo.docx"
doc = Document(ruta_plantilla)  

# Primera parte: Crear el diccionario de reemplazos

# Ejemplo 1: Reemplazo de texto en párrafos (dentro de párrafos o creación de párrafos nuevos)
dict_de_reemplazos = {
    "<<fecha>>": "03 de junio de 2026",
    "<<nombre_de_cliente>>": "BelloDev",
    "<<mes>>": "junio",
    "<<anio>>": "2026",
    "<<parrafos_previos>>": [
        ("Con el modulo de ezword se pueden crear parrafos con diferentes estilos.", "Normal"),
        ('Como por ejemplo, este párrafo con estilo "Negritas""', 'Negritas')]
}

# Ejemplo 2: insertar una lista
dict_de_reemplazos["<<lista_de_objetivos>>"] = [ 
                ("Primera recomendación","Normal"),
                ("Segunda recomendación","Normal")]


# Ejemplo 3: insertar una figura sin título
dict_de_reemplazos["<<figura_ejemplo>>"] = [{
        "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_1_1.jpg",
        "titulo": "", # No lleva título porque ya está en el documento
        "tamanio": 2,
        "bookmark": "", # No lleva bookmark porque ya está en el documento
        "estilo_figura":"Figura",
        "estilo_titulo": "Carcentrado"
    }]


dict_de_reemplazos["<<figura_con_titulo>>"] = [{
        "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_1_1.jpg",
        "titulo": "Se ve la carita de una niña", # Título personalizado para esta figura
        "tamanio": 2,
        "bookmark": "<<RefFigura_sec_1_1>>", 
        "estilo_figura":"Figura",
        "estilo_titulo": "Carcentrado"
    }]

# Ejemplo 5: Creación de secciones en el documento, un título, varios párrafos, figuras e imágenes
dict_de_reemplazos["<<sec_prueba>>"] = [
        ( "<<titulo_1>>", "subtitulo"),    
        ("",""),
        ("<<contenido_1>>","Normal"),
        ("",""),
        ("<<contenido_1_parte2>>", "Normal"),
        ("",""),
        ("De la <<reffigura_sec_1>> se muestran resultados interesantes.", "Normal"),
        ("",""),
        ("<<fig_sec_1>>", "figura"),
        ("<<titulo_2>>", "subtitulo"),
        ("",""),
        ("<<contenido_2>>", "Normal"),
        ("",""),
        ("De la <<reffigura_sec_2>> se muestran más resultados interesantes.", "Normal"),
        ("",""),
        ("<<fig_sec_2>>", "figura")]

# El diccionario de reemplazo busca la variable <<sec_prueba>> en el documento, y al encontrarla, 
# inserta una sección completa con el contenido definido en la lista asociada a esa variable. 
# Cada tupla en la lista representa un bloque de contenido, donde el primer elemento es el texto o marcador 
# a insertar, y el segundo elemento es el estilo a aplicar (por ejemplo, "subtitulo", "Normal", "figura", etc.).





nuevas_variables = {
    "<<titulo_1>>": "Resultados de la sección 1.",
    "<<contenido_1>>": "En esta sección se muestran los resultados obtenidos en el experimento 1.",
    "<<contenido_1_parte2>>": "Además, se observa que los resultados son consistentes con lo esperado.",
    "<<fig_sec_1>>": [{
            "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_1_1.jpg",
            "titulo": "Momardo 1-1",
            "tamanio": 2,
            "bookmark": "<<Ref_sec_1_1>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_1_2.jpg",
            "titulo": "Momardo 1-2",
            "tamanio": 2,
            "bookmark": "<<Ref_sec_1_2>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_1_3.jpg",
            "titulo": "Momardo 1-3",
            "tamanio": 2,
            "bookmark": "<<Ref_sec_1_3>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        }],
    "<<titulo_2>>": "Resultados de la sección 2.",
    "<<contenido_2>>": "En esta sección se muestran los resultados obtenidos en el experimento 2.",
    "<<fig_sec_2>>": [{
            "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_2_1.jpg",
            "titulo": "Momardo 2-1",
            "tamanio": 2,
            "bookmark": "<<Ref_sec_2_1>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_2_2.jpg",
            "titulo": "Momardo 2-2",
            "tamanio": 2,
            "bookmark": "<<Ref_sec_2_2>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": "C:\\programacion\\codigos_python\\ezsnake\\tests\\figuras_demo\\fig_sec_2_3.jpg",
            "titulo": "Momardo 2-3",
            "tamanio": 2,
            "bookmark": "<<Ref_sec_2_3>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        }],
}


tabla = {
    "encabezado_1": ["dato1", "dato2", "dato3"],
    "encabezado_2": ["otrodato1", "otrodato2", "otrodato3"]
}
df = pd.DataFrame(tabla)
estilo_de_tabla = ezw.EstilosTabla()
estilo_de_tabla.set_color_del_header(0, (36, 64, 97))  # Azul 5ta columna al fondo de colores de word
estilo_de_tabla.set_color_de_columna(0, (255, 228, 225))  # Rosa claro
# estilo_de_tabla.set_color_de_fila(0, (224, 255, 255))  # Cian claro

diccionario_de_tablas = {
    "<<nuevatabla_tabla1>>": {
                "tabla": df,
                "estilos_de_tabla": estilo_de_tabla.to_dict(),
                "titulo": "Resultados del análisis",
                "bookmark": "RefTabla_Resultados_1",
            }
}

ezw.reemplazar_texto_en_plantilla(doc, dict_de_reemplazos) # Esto pondría en el documento los contenidos del ejemplo 1, 2, 3 y 5.
# ezw.reemplazar_texto_en_plantilla(doc, nuevas_variables) # Esto pondría en el documento los contenidos del ejemplo 1 y 5.
# El ejemplo 5 es una sección, que introduce más variables nuevas al documento, como <<titulo_1>>, <<contenido_1>>, 
# <<fig_sec_1>>, etc. Estas variables también se reemplazarán en esta misma llamada a reemplazar_texto_en_plantilla,
# porque el código de esa función es capaz de detectar las nuevas variables que se introducen al insertar la sección, 
# y reemplazarlas también. Pero esto solo pasará si dentro de diccionario_de_reemplazos, existen esas variables.
# En este ejemplo el diccionario_de_reemplazos incluye las variables.

ezw.reemplazar_variable_por_figura(doc, dict_de_reemplazos)
ezw.reemplazar_referencias_cruzadas_de_figuras(doc, dict_de_reemplazos)
ezw.reemplazar_variable_por_tabla(doc, diccionario_de_tablas)


doc.save("C:\\programacion\\codigos_python\\ezsnake\\tests\\doc_listo.docx")
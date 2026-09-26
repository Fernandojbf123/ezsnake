import os
import ezsnake.word_template_writer as ezw
from docx import Document
import pandas as pd



ruta_plantilla =  os.path.join(os.getcwd(), "tests", "demo.docx")
doc = Document(ruta_plantilla)  

# Primera parte: Crear el diccionario de reemplazos
dict_de_reemplazos = {}

# Ejemplo 1: Reemplazo de texto en párrafos (dentro de párrafos o creación de párrafos nuevos)
dict_de_reemplazos["<<fecha>>"] = "03 de junio de 2026"

# Ejemplo 2: Reemplazo de varias variables respetando formatos y referencias.
dict_de_reemplazos["<<nombre_del_programador>>"] = "BelloDev"

# Ejemplo 3: Reemplazar una variable por todo un párrafo y asignarle un estilo a dicho párrafo. 
# Ahora, usemos una fecha calculada para el día de hoy.
fecha_hoy = pd.Timestamp.now().strftime("%d de %B de %Y") 
dict_de_reemplazos["<<parrafos_previos>>"] = [
    ('Algo interesante del módulo word_template_writer de ezsnake es que se pueden crear parrafos con diferentes estilos. Para asignarle estilos a un párrafo, se necesita tener el estilo guardado en la plantilla de Word. Por ejemplo, este párrafo tiene estilo "Normal".', "Normal"),
    ("","Normal"),
    (f'Este otro párrafo con estilo "Negritas" y usando una fecha calculada desde código: {fecha_hoy}', 'Negritas'),
    ("","Normal"),
    ('Este último párrafo tiene estilo "Normal_sin_sangria"', "Normal_sin_sangria")
]


# Ejemplo 4: insertar una lista
dict_de_reemplazos["<<lista_ejemplo>>"] = [ 
    "Primer punto de la lista.",
    "Segundo punto de la lista.",
    "Tercer punto de la lista.",
    "Puedes poner cuantos puntos quieras."
]


# Ejemplo 5: insertar una figura sin título
# El formato de código para generar una figura es como sigue:

dict_de_reemplazos["<<fig_sintitulo>>"] = [
    {
        "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_1_1.jpg"),
        "titulo": "", # No lleva título porque ya está en el documento
        "tamanio": 2, # El tamaño de la figura 
        "bookmark": "", # No lleva bookmark porque ya está en el documento (el bookmark es el título del a figura en word)
        "estilo_figura":"Figura",
        "estilo_titulo": "Carcentrado"
    }
]


# Ejemplo 6: insertar una figura con título
titulo_de_figura = "Robert Downie Jr. suspirando porque el titulo salió con estilo Carjustificado porque supera los 150 caracteres y todo está excelentemente bien" 
# titulo_de_figura = "Como el título es corto. Ahora debe salir con estilo Carcentrado"

estilo_titulo_de_figura = "Carjustificado" if len(titulo_de_figura) > 120 else "Carcentrado"

dict_de_reemplazos["<<fig_contitulo>>"] = [
    {
        "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_1_2.jpg"),
        "titulo": titulo_de_figura, # Título personalizado para esta figura
        "tamanio": 2,
        "bookmark": "<<RefFigura_contitulo>>", 
        "estilo_figura":"Figura",
        "estilo_titulo": estilo_titulo_de_figura
    }
]

# Ejemplo 7: Rellenar una tabla existente

# Ejemplo 7.1: Rellenar tabla y editar algunas propiedades de esta
estilos = ezw.EstilosTabla() # Crear la instancia de estilos; esto permite aplicar estilos del documento de word a la tabla
opciones = ezw.OpcionesTabla() # Crear la instancia de opciones; esto permite configurar opciones específicas para la tabla

# Creare una tabla de tres filas y tres columnas.
# El dataframe tiene los datos de la tabla
df = pd.DataFrame(
    {
        "Nombre": ["Fernando","José","Azeroth"], 
        "Apellido": ["Bello", "Fuentes", "Stormrage"],
        "Edad": [21, 23, 2500]
    }
)

# Supongamos que ahora; supongamos que la columna Edad debe ir en un estilo justificado; y que las otras dos columnas van en centrado
# Dentro del word existe el estilo texto_tabla_justificado
estilos.set_estilo_de_columna(0, 'tabla_texto_centrado') # La columna 0 es la columna Nombre
estilos.set_estilo_de_columna(1, 'tabla_texto_centrado') # La columna 1 es la columna Apellido
estilos.set_estilo_de_columna(2, 'tabla_texto_justificado') # La columna 2 es la columna Edad

# Nota: También se pueden asignar colores a una celda o a una columna entera; recomiendo leer el la ayuda de la clase EstilosTabla
estilos.set_color_de_celda(celda = (2,2), color = (75, 120, 200)) # acá estoy poniendo este color azul a la ultima celda
# Ahora las opciones; 
dict_de_reemplazos["<<editartabla_tabla_1>>"] = {
    "tabla": df,
    "estilos_de_tabla": estilos,
    "opciones_de_tabla": opciones,
}

# Ejemplo 7.2:Rellenar tabla agrupando valores.
# Ejemplo 7.1: Rellenar tabla y editar algunas propiedades de esta
estilos = ezw.EstilosTabla() # Crear la instancia de estilos; esto permite aplicar estilos del documento de word a la tabla
opciones = ezw.OpcionesTabla() # Crear la instancia de opciones; esto permite configurar opciones específicas para la tabla

# Creare una tabla de tres filas y tres columnas.

# El dataframe tiene los datos de la tabla
df = pd.DataFrame(
    {
        "Nombre": ["Fernando","Fernando","Fernando","José"], 
        "Apellido": ["Bello", "Fuentes", "Stormrage", "Bello"],
        "Edad": [21, 23, 2500, 40]
    }
)

# Ejemplo, combinar la primera columna
opciones.set_detectar_merge(detectar=True)  # Activar la detección de combinación para la tabla
opciones.set_columnas_para_merge(columnas=[0])  
dict_de_reemplazos["<<editartabla_tabla_2>>"] = {
    "tabla": df,
    "estilos_de_tabla": estilos,
    "opciones_de_tabla": opciones,
}

# Poner estilos de la tabla 2
estilos.set_estilo_de_columna(0, 'tabla_texto_centrado') # La columna 0 es la columna Nombre
estilos.set_estilo_de_columna(1, 'tabla_texto_centrado') # La columna 1 es la columna Apellido
estilos.set_estilo_de_columna(2, 'tabla_texto_centrado') # La columna 2 es la columna Edad


# Ejemplo 7.2: Crear una tabla desde cero; incluyendo título y referencias cruzadas.

## Quedé acá. Debo modificar esta parte para que avancemos














# Ejemplo 9: Agregar una sección
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
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_1_1.jpg"),
            "titulo": "Momardo 1-1",
            "tamanio": 2,
            "bookmark": "<<RefFigura_sec_1_1>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_1_2.jpg"),
            "titulo": "Momardo 1-2",
            "tamanio": 2,
            "bookmark": "<<RefFigura_sec_1_2>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_1_3.jpg"),
            "titulo": "Momardo 1-3",
            "tamanio": 2,
            "bookmark": "<<RefFigura_sec_1_3>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        }],
    "<<titulo_2>>": "Resultados de la sección 2.",
    "<<contenido_2>>": "En esta sección se muestran los resultados obtenidos en el experimento 2.",
    "<<fig_sec_2>>": [{
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_2_1.jpg"),
            "titulo": "Momardo 2-1",
            "tamanio": 2,
            "bookmark": "<<RefFigura_sec_2_1>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_2_2.jpg"),
            "titulo": "Momardo 2-2",
            "tamanio": 2,
            "bookmark": "<<RefFigura_sec_2_2>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        },
        {
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_sec_2_3.jpg"),
            "titulo": "Momardo 2-3",
            "tamanio": 2,
            "bookmark": "<<RefFigura_sec_2_3>>",
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

# ezw.reemplazar_texto_en_plantilla(doc, dict_de_reemplazos) # Esto pondría en el documento los contenidos del ejemplo 1, 2, 3 y 5.
# ezw.reemplazar_texto_en_plantilla(doc, nuevas_variables) # Esto pondría en el documento los contenidos del ejemplo 1 y 5.
# El ejemplo 5 es una sección, que introduce más variables nuevas al documento, como <<titulo_1>>, <<contenido_1>>, 
# <<fig_sec_1>>, etc. Estas variables también se reemplazarán en esta misma llamada a reemplazar_texto_en_plantilla,
# porque el código de esa función es capaz de detectar las nuevas variables que se introducen al insertar la sección, 
# y reemplazarlas también. Pero esto solo pasará si dentro de diccionario_de_reemplazos, existen esas variables.
# En este ejemplo el diccionario_de_reemplazos incluye las variables.

# Nota Se pueden correr todos los reemplazos de una sola vez si se unen los diccionarios:
# ejemplo
dict_unido = dict_de_reemplazos | nuevas_variables
ezw.reemplazar_texto_en_plantilla(doc, dict_de_reemplazos) # Esto pondría en el documento los contenidos del ejemplo 1 y 5.

ezw.reemplazar_variable_por_figura(doc, dict_de_reemplazos) # Pone las figuras sin título y con título.
ezw.reemplazar_referencias_cruzadas_de_figuras(doc, dict_de_reemplazos)

ezw.rellenar_tablas_en_plantilla(doc, dict_de_reemplazos)

# ezw.reemplazar_variable_por_tabla(doc, diccionario_de_tablas)

doc.save(os.path.join(os.getcwd(), "tests", "doc_listo.docx"))
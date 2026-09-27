import os
import ezsnake.word_template_writer as ezw
from docx import Document
import pandas as pd



ruta_plantilla =  os.path.join(os.getcwd(), "tests", "demo.docx")
doc = Document(ruta_plantilla)  

# Primera parte: Crear el diccionario de reemplazos
dict_de_reemplazos = {}


## Ejemplo 1: Reemplazo de texto en párrafos (dentro de párrafos o creación de párrafos nuevos)
dict_de_reemplazos["<<fecha>>"] = "03 de junio de 2026"


## Ejemplo 2: Reemplazo de varias variables respetando formatos y referencias.
dict_de_reemplazos["<<nombre_del_programador>>"] = "BelloDev"


## Ejemplo 3: Reemplazar una variable por todo un párrafo y asignarle un estilo a dicho párrafo. 
# Ahora, usemos una fecha calculada para el día de hoy.
fecha_hoy = pd.Timestamp.now().strftime("%d de %B de %Y") 
dict_de_reemplazos["<<parrafos_previos>>"] = [
    ('Algo interesante del módulo word_template_writer de ezsnake es que se pueden crear parrafos con diferentes estilos. Para asignarle estilos a un párrafo, se necesita tener el estilo guardado en la plantilla de Word. Por ejemplo, este párrafo tiene estilo "Normal".', "Normal"),
    ("","Normal"),
    (f'Este otro párrafo con estilo "Negritas" y usando una fecha calculada desde código: {fecha_hoy}', 'Negritas'),
    ("","Normal"),
    ('Este último párrafo tiene estilo "Normal_sin_sangria"', "Normal_sin_sangria")
]


## Ejemplo 4: insertar una lista
dict_de_reemplazos["<<lista_ejemplo>>"] = [ 
    "Primer punto de la lista.",
    "Segundo punto de la lista.",
    "Tercer punto de la lista.",
    "Puedes poner cuantos puntos quieras."
]


## Ejemplo 5: insertar una figura sin título
# El formato de código para generar una figura es como sigue:
dict_de_reemplazos["<<fig_sintitulo>>"] = [
    {
        "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_ejemplo_5.jpg"),
        "titulo": "", # No lleva título porque ya está en el documento
        "tamanio": 2, # El tamaño de la figura 
        "bookmark": "", # No lleva bookmark porque ya está en el documento (el bookmark es el título del a figura en word)
        "estilo_figura":"Figura",
        "estilo_titulo": "Carcentrado"
    }
]


## Ejemplo 6: insertar una figura con título
titulo_de_figura = "Una persona suspirando porque el titulo salió con estilo Carjustificado porque supera los 150 caracteres y todo está excelentemente bien" 
# titulo_de_figura = "Como el título es corto. Ahora debe salir con estilo Carcentrado"

estilo_titulo_de_figura = "Carjustificado" if len(titulo_de_figura) > 120 else "Carcentrado"

dict_de_reemplazos["<<fig_contitulo>>"] = [
    {
        "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "fig_ejemplo_6.jpg"),
        "titulo": titulo_de_figura, # Título personalizado para esta figura
        "tamanio": 2,
        "bookmark": "RefFigura_contitulo", 
        "estilo_figura":"Figura",
        "estilo_titulo": estilo_titulo_de_figura
    }
]

## Ejemplo 7: insertar varias figuras y poner una referencia hacia estas.

dict_de_reemplazos["<<fig_variasfiguras>>"] = [
     {
        "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "homero.jpg"),
        "titulo": "Este es Homero con formato Car_justificado.", # Título personalizado para esta figura
        "tamanio": 2,
        "bookmark": "RefFigura_variasfiguras_1", 
        "estilo_figura":"Figura",
        "estilo_titulo": "Carjustificado"
    },
    {
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", "pikachu.jpg"),
            "titulo": "Este es Pikachu con estilo Car_centrado.", # Título personalizado para esta figura
            "tamanio": 2,
            "bookmark": "RefFigura_variasfiguras_2", 
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
        }
]

## Ejemplo 8: Rellenar una tabla existente

### Ejemplo 8.1: Rellenar tabla y editar algunas propiedades de esta
# Paso 1: Crear instancias de estilos y opciones para la tabla.
estilos = ezw.EstilosTabla() # Crear la instancia de estilos; esto permite aplicar estilos del documento de word a la tabla
opciones = ezw.OpcionesTabla() # Crear la instancia de opciones; esto permite configurar opciones específicas para la tabla

# Paso 2: Creare una tabla de tres filas y tres columnas.
# El dataframe tiene los datos de la tabla
df = pd.DataFrame(
    {
        "Nombre": ["Fernando","José","Azeroth"], 
        "Apellido": ["Bello", "Fuentes", "Stormrage"],
        "Edad": [21, 23, 2500]
    }
)

# Paso 3: Aplicar estilos a las columnas de la tabla.
# Supongamos que ahora; supongamos que la columna Edad debe ir en un estilo justificado; y que las otras dos columnas van en centrado
# Dentro del word existe el estilo texto_tabla_justificado
estilos.set_estilo_de_columna(0, 'tablatextocentrado') # La columna 0 es la columna Nombre
estilos.set_estilo_de_columna(1, 'tablatextocentrado') # La columna 1 es la columna Apellido
estilos.set_estilo_de_columna(2, 'tablatextojustificado') # La columna 2 es la columna Edad

# Nota: También se pueden asignar colores a una celda o a una columna entera; recomiendo leer el la ayuda de la clase EstilosTabla
estilos.set_color_de_celda(celda = (2,2), color = (75, 120, 200)) # acá estoy poniendo este color azul a la ultima celda

# Paso 4: Asignar valores al diccionario de reemplazos
dict_de_reemplazos["<<editartabla_tabla_1>>"] = {
    "tabla": df,
    "estilos_de_tabla": estilos,
    "opciones_de_tabla": opciones,
}



### Ejemplo 8.2:Rellenar tabla agrupando valores.
# Paso 1: Crear instancias de estilos y opciones para la tabla.
estilos = ezw.EstilosTabla() # Crear la instancia de estilos; esto permite aplicar estilos del documento de word a la tabla
opciones = ezw.OpcionesTabla() # Crear la instancia de opciones; esto permite configurar opciones específicas para la tabla

# Paso 2: Creare una tabla de cuatro filas y tres columnas.
# El dataframe tiene los datos de la tabla
df = pd.DataFrame(
    {
        "Nombre": ["Fernando","Fernando","Fernando","José"], 
        "Disfraz": ["Alien", "Astronauta", "Pirata", "Científico"],
        "Puntaje": [100, 100, 30, 95]
    }
)

# Paso 3: poner las opciones; en este ejemplo combinar la primera columna
opciones.set_detectar_merge(detectar=True)  # Activar la detección de combinación para la tabla
opciones.set_columnas_para_merge(columnas=[0])  

# Paso 4: Colocar estilos a la tabla.
estilos.set_estilo_de_columna(0, 'tablatextocentrado') # La columna 0 es la columna Nombre
estilos.set_estilo_de_columna(1, 'tablatextocentrado') # La columna 1 es la columna Apellido
estilos.set_estilo_de_columna(2, 'tablatextocentrado') # La columna 2 es la columna Edad

# Paso 5: Asignar valores al diccionario de reemplazos
dict_de_reemplazos["<<editartabla_tabla_2>>"] = {
    "tabla": df,
    "estilos_de_tabla": estilos,
    "opciones_de_tabla": opciones,
}


## Ejemplo 9: Crear una tabla desde cero; incluyendo título y referencias cruzadas.

# Paso 1: Crear instancias de estilos y opciones para la tabla.
estilos = ezw.EstilosTabla() # Crear la instancia de estilos; esto permite aplicar estilos del documento de word a la tabla
opciones = ezw.OpcionesTabla() # Crear la instancia de opciones; esto permite configurar opciones específicas para la tabla

# Paso 2: Creare una tabla de tres filas y tres columnas.

# El dataframe tiene los datos de la tabla
df = pd.DataFrame(
    {
        "Ubicación": ["Av. Bolivar","Calle Miranda","Calzada Santiago Mariño"], 
        "Destino": ["Av. Libertador", "Av. Sucre", "Calle Negro Primero"],
        "Tiempo estimado": ["21 minutos", "23 minutos", "40 minutos"]
    }
)

# Paso 3: Asignar opciones: Ejemplo combinar la primera columna de la tabla.
opciones.set_detectar_merge(detectar=True)  # Activar la detección de combinación para la tabla
opciones.set_columnas_para_merge(columnas=[0])  

# Paso 4: Poner estilos a la nueva tabla creada desde cero
estilos.set_color_del_header(fila_header=0, color=(50, 100, 164))
estilos.set_estilo_del_header(fila_header=0, estilo='tablaheadercentradoblanco')

estilos.set_estilo_de_columna(0, 'tablatextocentrado') # La columna 0 es la columna Ubicación
estilos.set_estilo_de_columna(1, 'tablatextocentrado') # La columna 1 es la columna Destino
estilos.set_estilo_de_columna(2, 'tablatextocentrado') # La columna 2 es la columna Tiempo estimado

# Paso 5: Asignar valores al diccionario de reemplazos para la nueva tabla creada desde cero.
titulo= "Esta es una tabla creada desde cero."

dict_de_reemplazos["<<nuevatabla_tablacreadadesdecero>>"] = {
    "tabla": df,
    "estilos_de_tabla": estilos,
    "opciones_de_tabla": opciones,
    "titulo": titulo,
    "bookmark": "RefTabla_tablacreadadesdecero",
    "estilo_de_titulo": "Carcentrado"
}


# Ejemplo 10: Agregar una sección
# Primero vamos a leer el excel usando
df_entrevistas = pd.read_excel(os.path.join(os.getcwd(), "tests", "database", "base_de_datos_entrevistas.xlsx"))
df_entrevistas["fecha_de_entrevista"] = pd.to_datetime(df_entrevistas["fecha_de_entrevista"], format= "%d/%m/%y", errors='coerce')
fechas_unicas = df_entrevistas["fecha_de_entrevista"].dt.date.unique()
fechas_unicas_str = [fecha.strftime("%d/%m/%Y") for fecha in fechas_unicas]

def conector(lista: list) -> str:
    """ una pequeña función para poner "," o "y" entre los elementos de una lista"""
    texto = ""
    for elemento in lista:
        if len(lista) == 1:
            texto += str(elemento)
        else:
            if elemento == lista[-1]:
                texto += " y "
            else:
                texto += ", "
            
            texto += f"{elemento}"

    return texto.strip()

# Inicio con la sección totalmente vacía
seccion = []

# como pondré imágenes con títulos y quiero asegurar referencias únicas, usaré un contador para cada entrevistado.
contador_subtitulos = 0
contador_de_referencias_de_figuras = 0

# Creo un párrafo general que introduce la sección
parrafo_general = f"En esta sección se detallan las entrevistas realizadas en las fechas {conector(fechas_unicas_str)}."
seccion.append((parrafo_general, "Normal"))
seccion.append(("", ""))


# Ahora iniciamos un ciclo automatizado que se detalla cada día
dict_figuras_temporal = {}
for fecha in fechas_unicas:

    df_dia = df_entrevistas[df_entrevistas["fecha_de_entrevista"].dt.date == fecha]

    fecha_str = fecha.strftime("%d/%m/%Y")

    titulo = f"Entrevistas del {fecha_str}"
    nombres_de_entrevistados = df_dia["nombre"].tolist()
    compatibilidad = df_dia["compatibilidad"].tolist()
    cantidad_de_entrevistados = len(nombres_de_entrevistados)
    rutas_a_la_fotos = df_dia["ruta_a_la_foto"].tolist()
    lugar_de_entrevista = df_dia["lugar_de_entrevista"].unique().tolist()[0]
    realizada_por = df_dia["realizada_por"].unique().tolist()[0]

    # Creación del párrafo descriptivo según la cantidad de entrevistados
    if cantidad_de_entrevistados > 1:
        parrafo = f"El {fecha_str} se realizaron {cantidad_de_entrevistados} entrevistas en {lugar_de_entrevista}; las cuales estuvieron a cargo de {realizada_por}."
        parrafo += f" Los entrevistados fueron: {conector(nombres_de_entrevistados)}."
        parrafo += f" A continuación, de la <<reffigura_entrevistado_{contador_subtitulos}>> muestran las fotos de los entrevistados con su nombre y afinidad al puesto."
    
    if cantidad_de_entrevistados == 1:
        parrafo = f"El {fecha_str} se realizó una entrevista en {lugar_de_entrevista}; la cual estuvo a cargo de {realizada_por}."
        parrafo += f" El entrevistado fue: {conector(nombres_de_entrevistados)}."
        parrafo += f" A continuación, en la <<reffigura_entrevistado_{contador_subtitulos}>> se muestra la foto del entrevistado con su nombre y afinidad al puesto."
    
    seccion.append((titulo, "Negritas"))
    seccion.append(("", ""))
    seccion.append((parrafo, "Normal"))
    seccion.append(("", ""))


    # Ahora agreguemos varias figuras: Para esto debo crear una variable de figuras;
    seccion.append(f"<<fig_entrevistado_{contador_subtitulos}>>")
    # También debo crear un diccionario con el array de diccionarios de figuras. 
    # Al terminar el proceso de adjuntar figuras, este diccionario se unirá al dict_de_reemplazos
    dict_figuras_temporal[f"<<fig_entrevistado_{contador_subtitulos}>>"] = []
    
    for i_entrevistado in range(0, cantidad_de_entrevistados):
        dict_figuras_temporal[f"<<fig_entrevistado_{contador_subtitulos}>>"].append(
            {
            "ruta": os.path.join(os.getcwd(), "tests", "figuras_demo", f"{rutas_a_la_fotos[i_entrevistado]}"),
            "titulo": f"{nombres_de_entrevistados[i_entrevistado]}, compatibilidad = {compatibilidad[i_entrevistado]}",
            "tamanio": 2,
            "bookmark": f"<<RefFigura_entrevistado_{contador_de_referencias_de_figuras}>>",
            "estilo_figura":"Figura",
            "estilo_titulo": "Carcentrado"
            }
        )
        contador_de_referencias_de_figuras += 1

    
    contador_subtitulos += 1


# Ahora a la variable le agrego todo el contenido de sección
dict_de_reemplazos["<<sec_ejemplo>>"] = seccion
# Y ahora agrego el diccionario de figuras temporales al diccionario de reemplazos
dict_de_reemplazos = dict_de_reemplazos | dict_figuras_temporal



# Parte final: Llamar los scripts que se encargan de reemplazar el texto, 
# las figuras, las referencias cruzadas y las tablas en la plantilla.
ezw.reemplazar_texto_en_plantilla(doc, dict_de_reemplazos) # Esto pondría en el documento los contenidos del ejemplo 1 y 5.
ezw.reemplazar_variable_por_figura(doc, dict_de_reemplazos) # Pone las figuras sin título y con título.
ezw.reemplazar_referencias_cruzadas_de_figuras(doc, dict_de_reemplazos)
ezw.rellenar_tablas_en_plantilla(doc, dict_de_reemplazos)
ezw.reemplazar_variable_por_tabla(doc, dict_de_reemplazos)
ezw.reemplazar_referencias_cruzadas_de_tablas(doc, dict_de_reemplazos)

# Guardar documento
doc.save(os.path.join(os.getcwd(), "tests", "doc_listo.docx"))
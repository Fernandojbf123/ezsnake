import os
import pandas as pd
from configs.manager_doc_config import *

# modulo de construccion de diccionarios para templates de word de ezsnake by BelloDev
from services.word_template_writer import *


# Managers de variables de excel
from services.manager_variables_excel_datos_campania import *
from services.manager_variables_excel_datos_despliegue import *
from services.manager_variables_excel_documento import *

# Constructores individuales
from services.doc_101.descripcion_actividades_previas.descripcion_actividades_previas_parrafo1 import *
from services.doc_101.descripcion_actividades_previas.descripcion_actividades_previas_parrafo2 import *
from services.doc_101.descripcion_actividades_previas.descripcion_actividades_previas_parrafo3 import *

from services.doc_101.ejecucion_de_la_campania.ejecucion_de_la_campania_parrafo1 import *
from services.doc_101.ejecucion_de_la_campania.ejecucion_de_la_campania_parrafo2 import *
from services.doc_101.ejecucion_de_la_campania.ejecucion_de_la_campania_parrafo3 import *
from services.doc_101.ejecucion_de_la_campania.ejecucion_de_la_campania_parrafo4 import *


############################ DICCIONARIO DE REEMPLAZOS PARA FIGURAS ############################

# ESQUEMA DEL DICCIONARIO DE FIGURAS
class Dictfiguras():
    def __init__(self):
        self.ruta = ""
        self.titulo = ""
        self.tamanio = 6
        self.bookmark = ""

    def set_ruta(self, varvalue: str):
        carpeta = get_ruta_a_carpeta_de_las_figuras(usar_NAS=True)
        ruta_completa = os.path.join(carpeta, varvalue+".jpg")
        self.ruta = ruta_completa.strip()
    
    def set_tamanio(self, tamanio: int):
        self.tamanio = tamanio
    
    def set_bookmark(self, varvalue: str):
        bookmark = "Ref_"+varvalue
        self.bookmark = bookmark.strip()
        
    def set_titulo(self, varvalue: str):
        titulo = varvalue.strip() 
        if titulo != "":
            titulo = titulo if titulo.endswith(".") else titulo + "."
        self.titulo = titulo
        
    def return_dict(self) -> dict:
        return {
            "ruta": self.ruta,
            "titulo": self.titulo,
            "tamanio": self.tamanio,
            "bookmark": self.bookmark
        }

def construir_diccionario_agregar_figuras(df_datos_documento: pd.DataFrame) -> dict:
    dict_documento = {}
    varnames_documento = get_varnames_documento(df_datos_documento = df_datos_documento) 

    for ivarname, varname in enumerate(varnames_documento):
        varvalues = get_variable_documento(df_datos_documento = df_datos_documento, nombre_variable = varname)
        if varname.startswith("fig_"):
            dict_temporal = Dictfiguras()  
            array=[]
            for ivarvalue, varvalue in enumerate(varvalues):
                dict_temporal.set_ruta(varvalue)
                dict_temporal.set_tamanio(3) 
                dict_temporal.set_bookmark(varvalue)
                dict_temporal.set_titulo("")
                next_varname = varnames_documento[ivarname+1]
                
                if next_varname.startswith("pie_"):
                    pie_value = get_variable_documento(df_datos_documento = df_datos_documento, nombre_variable = next_varname)
                    dict_temporal.set_titulo(pie_value[ivarvalue])
                
                if varname.lower() == "fig_mapa_de_despliegue".lower():
                    dict_temporal.set_tamanio(6)  

                if varname.lower() == "fig_esquema_de_sonda".lower():
                    numero_de_serie = varvalue.split("_")[-1]
                    titulo = f"Despliegue de sonda oceanográfica {numero_de_serie}"
                    dict_temporal.set_titulo(titulo)
                    dict_temporal.set_tamanio(6)  
                    
                elif varname.lower() == "fig_pruebas_de_transmision".lower():
                    numero_de_serie = varvalue.split("_")[-1]
                    titulo = f"Datos enviados durante las pruebas de laboratorio para la sonda {numero_de_serie}"
                    dict_temporal.set_titulo(titulo)
                    dict_temporal.set_tamanio(6)
                    
                elif varname.lower() == "fig_pruebas_baterias".lower():
                    titulo = f"Datos transmitidos del estado de las baterias durante las 24 horas de las pruebas de funcionamiento"
                    dict_temporal.set_titulo(titulo)
                    dict_temporal.set_tamanio(6)  
                
                elif varname.lower() == "fig_ubicacion_durante_pruebas".lower():
                    titulo = f"Mapa con la información con las primeras 24 horas de transmisión de las sondas"
                    dict_temporal.set_titulo(titulo)
                    dict_temporal.set_tamanio(6)  
                    
                elif varname.lower() == "fig_pruebas_de_funcionamiento".lower():
                    dict_temporal.set_tamanio(6)

                array.append(dict_temporal.return_dict())
                
            dict_documento["<<"+varname+">>"] = array
            
    return dict_documento


############################ DICCIONARIO DE REEMPLAZOS PARA TEXTO ############################
def construir_diccionario_de_datos_documento(df_datos_despliegue: pd.DataFrame, 
                                            df_datos_campanias: pd.DataFrame,
                                            df_datos_documento: pd.DataFrame,
                                            diccionario_de_reemplazos: dict):
    
    ## orden de servicio
    diccionario_de_reemplazos["<<orden_de_servicio>>"] = get_orden_de_servicio()    
    
    ## fechas de vigencia
    diccionario_de_reemplazos["<<fecha_inicio_vigencia>>"] = get_fecha_inicio_vigencia(df_datos_despliegue = df_datos_despliegue)
    diccionario_de_reemplazos["<<fecha_final_vigencia>>"] = get_fecha_final_vigencia(df_datos_despliegue = df_datos_despliegue)
    
    # fecha de entrega
    diccionario_de_reemplazos["<<fecha_de_entrega>>"] = get_fecha_entrega(df_datos_despliegue = df_datos_despliegue)
    
    # seriales de sondas
    seriales_de_sondas = get_seriales_de_sondas(df_datos_despliegue = df_datos_despliegue)
    diccionario_de_reemplazos["<<seriales_de_sondas>>"] = ", ".join([str(serial) for serial in seriales_de_sondas])  # Convierte a string con formato "12345, 67890"
    diccionario_de_reemplazos["<<numero_de_sondas>>"] = get_numero_de_sondas(df_datos_despliegue = df_datos_despliegue) 
    
    ## mes y_año de liberacion
    diccionario_de_reemplazos["<<mes_y_anio_de_liberacion>>"] = get_mes_y_anio_de_liberacion(df_datos_campanias = df_datos_campanias)
    
    diccionario_de_reemplazos["<<descripcion_actividades_previas_parrafo1>>"] = descripcion_actividades_previas_parrafo1(df_datos_campanias = df_datos_campanias, 
                                                                                                                        df_datos_despliegue= df_datos_despliegue)
    
    diccionario_de_reemplazos["<<descripcion_actividades_previas_parrafo2>>"] = descripcion_actividades_previas_parrafo2(df_datos_campanias = df_datos_campanias, 
                                                                                                                        df_datos_despliegue= df_datos_despliegue)
    
    diccionario_de_reemplazos["<<descrpcion_actividades_previas_parrafo3>>"] = descripcion_actividades_previas_parrafo3(df_datos_campanias = df_datos_campanias, 
                                                                                                                        df_datos_despliegue= df_datos_despliegue)
    
    diccionario_de_reemplazos["<<ejecucion_de_la_campania_parrafo1>>"] = ejecucion_de_la_campania_parrafo1(df_datos_campanias = df_datos_campanias, 
                                                                                                            df_datos_despliegue= df_datos_despliegue)
    
    diccionario_de_reemplazos["<<ejecucion_de_la_campania_parrafo2>>"] = ejecucion_de_la_campania_parrafo2(df_datos_campanias = df_datos_campanias, 
                                                                                                            df_datos_despliegue= df_datos_despliegue)
    
    diccionario_de_reemplazos["<<ejecucion_de_la_campania_parrafo3>>"] = ejecucion_de_la_campania_parrafo3(df_datos_campanias = df_datos_campanias, 
                                                                                                            df_datos_despliegue= df_datos_despliegue)
    
    diccionario_de_reemplazos["<<ejecucion_de_la_campania_parrafo4>>"] = ejecucion_de_la_campania_parrafo4(df_datos_campanias = df_datos_campanias, 
                                                                                                            df_datos_despliegue= df_datos_despliegue)
    


############################ DICCIONARIO DE REEMPLAZOS PARA PLANES DE CRUCEROS ############################
def construir_diccionario_de_reemplazos_para_plan_de_cruceros(df_datos_despliegue: pd.DataFrame, 
                                                            df_datos_campanias: pd.DataFrame,
                                                            df_datos_documento: pd.DataFrame,
                                                            diccionario_de_reemplazos: dict):

    df_unicos= get_fecha_y_hora_de_embarque_y_campania_unicos(df_datos_campanias = df_datos_campanias)
    campanias_unicas = df_unicos["campania"].unique()
    rutas = []
    for campania in campanias_unicas:
        archivo = campania + ".docx"
        ruta = os.path.join(get_ruta_a_carpeta_de_planes_de_crucero(usar_NAS=True), archivo)
        rutas.append(ruta)
    diccionario_de_reemplazos["<<external_doc_plan_de_crucero>>"] = rutas

    
############################# DICCIONARIO DE REEMPLAZOS PARA TABLAS ############################
# Es probable que acá necesite varios esquemas, dependiendo de la tabla.
def construir_diccionario_de_reemplazos_para_tablas(df_datos_despliegue: pd.DataFrame, 
                                                    df_datos_campanias: pd.DataFrame,
                                                    df_datos_documento: pd.DataFrame,
                                                    diccionario_de_reemplazos: dict,
                                                    doc: object):
    
    
    opciones_de_tabla = OpcionesTabla()
    estilos_de_tabla = EstilosTabla(doc)
    estilos_de_tabla.set_estilo_por_defecto("texto_tablas_centrado")
    tabla2 = df_datos_despliegue[["serial_de_sonda","latitud_maniobra","longitud_maniobra"]]
    tabla2.insert(0,"secuencia", range(1, len(tabla2) + 1))
    tabla2["secuencia"] = tabla2["secuencia"].astype(int).astype(str)
    tabla2["serial_de_sonda"] = tabla2["serial_de_sonda"].astype(int).astype(str)
    tabla2["latitud_maniobra"] = tabla2["latitud_maniobra"].astype(str)
    tabla2["longitud_maniobra"] = tabla2["longitud_maniobra"].astype(str)
    diccionario_de_reemplazos["<<tabla_plan>>"] = {
        "tabla": tabla2,
        "estilos_de_tabla": estilos_de_tabla,
        "opciones_de_tabla": opciones_de_tabla
    }
    
    
    tabla3 = df_datos_despliegue[["serial_de_sonda","latitud_plan","longitud_plan","fecha_y_hora_de_despliegue_maniobra","estado_despliegue"]]
    tabla3.insert(0,"secuencia", range(1, len(tabla3) + 1))
    tabla3["secuencia"] = tabla3["secuencia"].astype(int).astype(str)
    tabla3["serial_de_sonda"] = tabla3["serial_de_sonda"].astype(int).astype(str)
    tabla3["latitud_plan"] = tabla3["latitud_plan"].astype(str)
    tabla3["longitud_plan"] = tabla3["longitud_plan"].astype(str)
    tabla3["fecha_y_hora_de_despliegue_maniobra"] = pd.to_datetime(tabla3["fecha_y_hora_de_despliegue_maniobra"], format = "%d/%m/%Y %H:%M:%S").dt.strftime("%d/%m/%Y %H:%M")
    diccionario_de_reemplazos["<<tabla_maniobra>>"] = {
        "tabla": tabla3,
        "estilos_de_tabla": estilos_de_tabla,
        "opciones_de_tabla": opciones_de_tabla
    }



    
    df_datos_despliegue["serial_de_sonda"] = df_datos_despliegue["serial_de_sonda"].astype(int).astype(str)
    
    equipos = ["GPS primario",
                "GPS secundario", 
                "Sistema de telemetría primario",
                "Sistema de telemetría secundario", 
                "Sensor de temperatura primario",
                "Sensor de temperatura secundario", 
                "Acelerómetro"]    

    seriales = get_seriales_de_sondas(df_datos_despliegue = df_datos_despliegue)
    
    array_secuencia_tabla4 = []
    array_seriales_tabla4 = []
    array_equipos_tabla4 = []
    array_numero_de_serie_de_equipo_tabla4 = []
    
    for secuencia, serial in enumerate(seriales) :
        array_seriales_tabla4.append([serial]*len(equipos))
        array_equipos_tabla4.append(equipos)
        gps_primario = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["gps_primario"].iloc[0]
        gps_secundario = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["gps_secundario"].iloc[0]
        telemetria_primario = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["telemetria_primario"].iloc[0]
        telemetria_secundario = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["telemetria_secundario"].iloc[0]
        temperatura_primario = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["temperatura_primario"].iloc[0]
        temperatura_secundario = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["temperatura_secundario"].iloc[0]
        acelerometro = df_datos_despliegue[df_datos_despliegue["serial_de_sonda"] == serial]["acelerometro"].iloc[0]    
        array_numero_de_serie_de_equipo_tabla4.append([gps_primario, gps_secundario, telemetria_primario, telemetria_secundario, temperatura_primario, temperatura_secundario, acelerometro])
        array_secuencia_tabla4.append([secuencia+1]*len(equipos))
    
    array_secuencia_tabla4 = np.array(array_secuencia_tabla4).flatten()
    array_seriales_tabla4 = np.array(array_seriales_tabla4).flatten()
    array_equipos_tabla4 = np.array(array_equipos_tabla4).flatten()
    array_numero_de_serie_de_equipo_tabla4 = np.array(array_numero_de_serie_de_equipo_tabla4).flatten()
    
    tabla4_dict = {
        "secuencia": array_secuencia_tabla4,
        "serial_de_sonda": array_seriales_tabla4,
        "equipo": array_equipos_tabla4,
        "numero_de_serie_de_equipo": array_numero_de_serie_de_equipo_tabla4
    }
    tabla4 = pd.DataFrame(tabla4_dict)
        
    opciones_de_tabla.set_detectar_merge(True)
    opciones_de_tabla.set_columnas_para_merge([0,1])
    estilos_de_tabla.set_estilo_de_columna(2, "texto_tablas_justificado")
    estilos_de_tabla.set_estilo_de_columna(3, "texto_tablas_justificado")   
    
    diccionario_de_reemplazos["<<tabla_equipos>>"] = {
        "tabla": tabla4,
        "estilos_de_tabla": estilos_de_tabla,
        "opciones_de_tabla": opciones_de_tabla
    }

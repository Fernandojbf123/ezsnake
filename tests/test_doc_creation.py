import ezsnake.word_template_writer as ezw
from docx import Document


ruta_plantilla =  "C:\\programacion\\codigos_python\\ezsnake\\tests\\demo.docx"
doc = Document(ruta_plantilla)  

dict_de_reemplazos = {
    "<<sec_prueba>>": [
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

}

ezw.reemplazar_texto_en_plantilla(doc, dict_de_reemplazos)

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

ezw.reemplazar_texto_en_plantilla(doc, nuevas_variables)
ezw.reemplazar_variable_por_figura(doc, nuevas_variables)
ezw.reemplazar_referencias_cruzadas_de_figuras(doc, nuevas_variables)

doc.save("C:\\programacion\\codigos_python\\ezsnake\\tests\\doc_listo.docx")
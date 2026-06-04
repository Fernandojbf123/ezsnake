"""
Schema Helpers for Word Template Writer
=======================================

This module provides configuration classes for table styling, table options,
and figure configuration with validation and convenient setter methods.

Classes:
    - EstilosTabla: Configuration for table cell styling (colors, alignment, styles)
    - OpcionesTabla: Configuration for table behavior (merge cells, flatten MultiIndex)
    - FigSchema: Configuration for individual figures
    - TablaSchema: Configuration for individual dynamic tables
    
Functions:
    - get_estilos_disponibles: Extract available style names from a Word document
"""

from typing import Optional, Union, List, Tuple
import os
import pandas as pd


def get_estilos_disponibles(doc) -> List[str]:
    """
    Extrae los nombres de todos los estilos de párrafo disponibles en un documento Word.
    
    Args:
        doc: Objeto Document de python-docx
    
    Returns:
        Lista de nombres de estilos de párrafo disponibles
    
    Example:
        >>> from docx import Document
        >>> doc = Document('plantilla.docx')
        >>> estilos = get_estilos_disponibles(doc)
        >>> print(estilos)
        ['Normal', 'Heading 1', 'texto_tablas_centrado', ...]
    """
    estilos = []
    for style in doc.styles:
        if style.type == 1:  # 1 = PARAGRAPH (WD_STYLE_TYPE.PARAGRAPH)
            estilos.append(style.name)
    return estilos


class EstilosTabla:
    """
    Clase para configurar estilos de celdas en tablas de Word.
    
    Permite configurar estilos por defecto, por columna, por fila y por celda específica.
    La jerarquía de aplicación es: celda > fila > columna > defecto.

        Nota de uso:
                - Por defecto, todas las filas y columnas usan la configuración de `por_defecto`.
                - Para modificar el header (fila 0), se recomienda:
                    * set_color_del_fondo_de_fila(0, (R, G, B))
                    * set_color_del_texto_de_fila(0, (R, G, B))
                    * set_estilo_de_fila(0, "estilo")
    
    Attributes:
        _doc: Documento Word para validación de estilos
        _estilos_disponibles: Lista de estilos disponibles en el documento
        _config: Diccionario interno con la configuración
    
    Example:
        >>> from docx import Document
        >>> doc = Document('plantilla.docx')
        >>> estilos = EstilosTabla(doc)
        >>> estilos.set_color_de_columna(1, (230, 230, 250))
        >>> estilos.set_estilo_de_columna(1, 'texto_tablas_justificado')
        >>> config_dict = estilos.to_dict()
    """
    
    def __init__(self):
        """
        Inicializa EstilosTabla con valores por defecto.
        
        Args:
            doc: Objeto Document de python-docx para validar estilos
        """
        self._config = {}
        self.set_default()  # Aplica valores por defecto automáticamente
    
    # ===== MÉTODOS PRINCIPALES =====
    
    def set_default(self):
        """Configura valores por defecto razonables para tablas"""
        self._config = {
            "por_defecto": {
                "estilo_parrafo": "Normal",
                "alineacion_vertical": "center",
                "altura_fila": 288290,  # ~4 cm en twips
                "color_fondo": None
            },
            "por_columna": {},
            "por_fila": {},
            "por_celda": {}
        }
    
    def reset(self):
        """Limpia toda la configuración (diccionario vacío)"""
        self._config = {
            "por_defecto": {},
            "por_columna": {},
            "por_fila": {},
            "por_celda": {}
        }
    
    def to_dict(self) -> dict:
        """
        Retorna el diccionario interno de configuración.
        
        Returns:
            Copia del diccionario de configuración
        """
        return self._config.copy()
    
    # ===== SETTERS POR DEFECTO =====
    
    def set_estilo_por_defecto(self, estilo: str):
        """
        Establece el estilo de párrafo por defecto.
        
        Args:
            estilo: Nombre del estilo de párrafo
        
        Raises:
            ValueError: Si el estilo no existe en el documento
        """
        self._config["por_defecto"]["estilo_parrafo"] = estilo
    
    def set_color_fondo_por_defecto(self, color: Optional[Tuple[int, int, int]]):
        """
        Establece el color de fondo por defecto (RGB).
        
        Args:
            color: Tupla (R, G, B) con valores 0-255, o None para sin color
        
        Raises:
            ValueError: Si el color no es válido
        """
        self._validar_color(color)
        self._config["por_defecto"]["color_fondo"] = color
    
    def set_alineacion_vertical_por_defecto(self, alineacion: str):
        """
        Establece la alineación vertical por defecto.
        
        Args:
            alineacion: 'center', 'top', o 'bottom'
        """
        self._config["por_defecto"]["alineacion_vertical"] = alineacion
    
    def set_altura_fila_por_defecto(self, altura: int):
        """
        Establece la altura de fila por defecto.
        
        Args:
            altura: Altura en twips (288290 ≈ 4cm)
        """
        self._config["por_defecto"]["altura_fila"] = altura
    
    # ===== SETTERS POR COLUMNA =====
    
    def set_estilo_de_columna(self, col: int, estilo: str):
        """
        Establece el estilo de párrafo para una columna.
        
        Args:
            col: Índice de la columna (0-based)
            estilo: Nombre del estilo de párrafo
        
        Raises:
            ValueError: Si col es negativo o si el estilo no existe
        """
        self._validar_indice_columna(col)
        if col not in self._config["por_columna"]:
            self._config["por_columna"][col] = {}
        self._config["por_columna"][col]["estilo_parrafo"] = estilo
    
    def set_color_de_columna(self, col: int, color: Tuple[int, int, int]):
        """
        Establece el color de fondo para una columna.
        
        Args:
            col: Índice de la columna (0-based)
            color: Tupla (R, G, B) con valores 0-255
        
        Raises:
            ValueError: Si col es negativo o si el color no es válido
        """
        self._validar_indice_columna(col)
        self._validar_color(color)
        if col not in self._config["por_columna"]:
            self._config["por_columna"][col] = {}
        self._config["por_columna"][col]["color_fondo"] = color
    
    def set_color_de_texto_de_columna(self, col: int, color: Tuple[int, int, int]):
        """
        Establece el color del texto para una columna.
        
        Args:
            col: Índice de la columna (0-based)
            color: Tupla (R, G, B) con valores 0-255
        
        Raises:
            ValueError: Si col es negativo o si el color no es válido
        """
        self._validar_indice_columna(col)
        self._validar_color(color)
        if col not in self._config["por_columna"]:
            self._config["por_columna"][col] = {}
        self._config["por_columna"][col]["color_texto"] = color
    
    def set_alineacion_horizontal_de_columna(self, col: int, alineacion: str):
        """
        Establece la alineación horizontal para una columna.
        
        Args:
            col: Índice de la columna (0-based)
            alineacion: 'left', 'center', 'right', o 'justify'
        
        Raises:
            ValueError: Si col es negativo
        """
        self._validar_indice_columna(col)
        if col not in self._config["por_columna"]:
            self._config["por_columna"][col] = {}
        self._config["por_columna"][col]["alineacion_horizontal"] = alineacion
    
    # ===== SETTERS POR FILA =====
    
    def set_color_de_fila(self, fila: int, color: Tuple[int, int, int]):
        """
        Establece el color de fondo para una fila.
        
        Args:
            fila: Índice de la fila (0-based)
            color: Tupla (R, G, B) con valores 0-255
        
        Raises:
            ValueError: Si fila es negativo o si el color no es válido
        """
        self._validar_indice_fila(fila)
        self._validar_color(color)
        if fila not in self._config["por_fila"]:
            self._config["por_fila"][fila] = {}
        self._config["por_fila"][fila]["color_fondo"] = color
    
    def set_estilo_de_fila(self, fila: int, estilo: str):
        """
        Establece el estilo de párrafo para una fila.
        
        Args:
            fila: Índice de la fila (0-based)
            estilo: Nombre del estilo de párrafo
        
        Raises:
            ValueError: Si fila es negativo o si el estilo no existe
        """
        self._validar_indice_fila(fila)
        if fila not in self._config["por_fila"]:
            self._config["por_fila"][fila] = {}
        self._config["por_fila"][fila]["estilo_parrafo"] = estilo

    def set_color_del_fondo_de_fila(self, fila: int, color: Tuple[int, int, int]):
        """Alias explícito para set_color_de_fila()."""
        self.set_color_de_fila(fila, color)

    def set_color_del_texto_de_fila(self, fila: int, color: Tuple[int, int, int]):
        """Establece color de texto para una fila (útil para fila header)."""
        self._validar_indice_fila(fila)
        self._validar_color(color)
        if fila not in self._config["por_fila"]:
            self._config["por_fila"][fila] = {}
        self._config["por_fila"][fila]["color_texto"] = color
    
    # ===== SETTERS POR CELDA =====
    
    def set_color_de_celda(self, celda: Tuple[int, int], color: Tuple[int, int, int]):
        """
        Establece el color de fondo para una celda específica.
        
        Args:
            celda: Tupla (fila, columna) con índices 0-based
            color: Tupla (R, G, B) con valores 0-255
        
        Raises:
            ValueError: Si la celda no es válida o si el color no es válido
        """
        self._validar_celda(celda)
        self._validar_color(color)
        if celda not in self._config["por_celda"]:
            self._config["por_celda"][celda] = {}
        self._config["por_celda"][celda]["color_fondo"] = color
    
    def set_estilo_de_celda(self, celda: Tuple[int, int], estilo: str):
        """
        Establece el estilo de párrafo para una celda específica.
        
        Args:
            celda: Tupla (fila, columna) con índices 0-based
            estilo: Nombre del estilo de párrafo
        
        Raises:
            ValueError: Si la celda no es válida o si el estilo no existe
        """
        self._validar_celda(celda)
        if celda not in self._config["por_celda"]:
            self._config["por_celda"][celda] = {}
        self._config["por_celda"][celda]["estilo_parrafo"] = estilo
    
    # ===== GETTERS =====
    
    def get_estilo_por_defecto(self) -> Optional[str]:
        """Retorna el estilo de párrafo por defecto"""
        return self._config["por_defecto"].get("estilo_parrafo")
    
    def get_color_de_columna(self, col: int) -> Optional[Tuple[int, int, int]]:
        """Retorna el color de fondo de una columna"""
        return self._config["por_columna"].get(col, {}).get("color_fondo")
    
    def get_config_de_columna(self, col: int) -> dict:
        """Retorna toda la configuración de una columna"""
        return self._config["por_columna"].get(col, {})
    
    def get_config_de_fila(self, fila: int) -> dict:
        """Retorna toda la configuración de una fila"""
        return self._config["por_fila"].get(fila, {})
    
    def get_config_de_celda(self, celda: Tuple[int, int]) -> dict:
        """Retorna toda la configuración de una celda"""
        return self._config["por_celda"].get(celda, {})
    
    # ===== VALIDADORES PRIVADOS =====
    
    def _validar_color(self, color: Optional[Tuple[int, int, int]]):
        """Valida que sea tupla de 3 enteros entre 0-255 o None"""
        if color is None:
            return
        if not isinstance(color, tuple) or len(color) != 3:
            raise ValueError(
                f"Color debe ser tupla de 3 valores (R, G, B) o None. Recibido: {color}"
            )
        if not all(isinstance(c, int) and 0 <= c <= 255 for c in color):
            raise ValueError(
                f"Valores RGB deben ser enteros entre 0-255. Recibido: {color}"
            )
    
    def _validar_indice_columna(self, col: int):
        """Valida que el índice de columna sea no negativo"""
        if not isinstance(col, int) or col < 0:
            raise ValueError(
                f"Índice de columna debe ser entero no negativo. Recibido: {col}"
            )
    
    def _validar_indice_fila(self, fila: int):
        """Valida que el índice de fila sea no negativo"""
        if not isinstance(fila, int) or fila < 0:
            raise ValueError(
                f"Índice de fila debe ser entero no negativo. Recibido: {fila}"
            )
    
    def _validar_celda(self, celda: Tuple[int, int]):
        """Valida que celda sea tupla (fila, columna) válida"""
        if not isinstance(celda, tuple) or len(celda) != 2:
            raise ValueError(
                f"Celda debe ser tupla (fila, columna). Recibido: {celda}"
            )
        self._validar_indice_fila(celda[0])
        self._validar_indice_columna(celda[1])


class OpcionesTabla:
    """
    Clase para configurar opciones de comportamiento de tablas.
    
    Controla aspectos como el aplanamiento de MultiIndex, detección de merged cells,
    y eliminación de filas marcadoras.
    
    Attributes:
        _opciones: Diccionario interno con las opciones
    
    Example:
        >>> opciones = OpcionesTabla()
        >>> opciones.set_detectar_merge(True)
        >>> opciones.set_columnas_para_merge([0, 1])
        >>> opciones_dict = opciones.to_dict()
    """
    
    def __init__(self):
        """Inicializa OpcionesTabla con valores por defecto"""
        self._opciones = {}
        self.set_default()  # Aplica valores por defecto automáticamente
    
    # ===== MÉTODOS PRINCIPALES =====
    
    def set_default(self):
        """Configura opciones por defecto razonables"""
        self._opciones = {
            "aplanar_multiindex": True,
            "detectar_merge": False,
            "columnas_para_merge": None,
            "eliminar_fila_marcador": True
        }
    
    def reset(self):
        """Limpia todas las opciones"""
        self._opciones = {}
    
    def to_dict(self) -> dict:
        """
        Retorna el diccionario interno de opciones.
        
        Returns:
            Copia del diccionario de opciones
        """
        return self._opciones.copy()
    
    # ===== SETTERS =====
    
    def set_aplanar_multiindex(self, aplanar: bool):
        """
        Si True, aplica reset_index() automáticamente a DataFrames con MultiIndex.
        Si False y se detecta MultiIndex, lanza error.
        
        Args:
            aplanar: Boolean para activar/desactivar aplanamiento
        
        Raises:
            ValueError: Si aplanar no es booleano
        """
        if not isinstance(aplanar, bool):
            raise ValueError(f"aplanar debe ser booleano. Recibido: {aplanar}")
        self._opciones["aplanar_multiindex"] = aplanar
    
    def set_detectar_merge(self, detectar: bool):
        """
        Si True, detecta y combina celdas verticales con valores consecutivos repetidos.
        Si False, no combina celdas (más rápido).
        
        Args:
            detectar: Boolean para activar/desactivar detección de merge
        
        Raises:
            ValueError: Si detectar no es booleano
        """
        if not isinstance(detectar, bool):
            raise ValueError(f"detectar debe ser booleano. Recibido: {detectar}")
        self._opciones["detectar_merge"] = detectar
    
    def set_columnas_para_merge(self, columnas: Optional[List[int]]):
        """
        Lista de índices de columnas donde aplicar merge vertical.
        None = todas las columnas, [] = ninguna, [0, 2] = solo columnas 0 y 2
        
        Args:
            columnas: Lista de índices de columnas o None
        
        Raises:
            ValueError: Si columnas no es lista válida o None
        """
        if columnas is not None:
            if not isinstance(columnas, list):
                raise ValueError(
                    f"columnas_para_merge debe ser lista o None. Recibido: {columnas}"
                )
            if not all(isinstance(c, int) and c >= 0 for c in columnas):
                raise ValueError(
                    f"Todos los índices deben ser enteros no negativos. Recibido: {columnas}"
                )
        self._opciones["columnas_para_merge"] = columnas
    
    def set_eliminar_fila_marcador(self, eliminar: bool):
        """
        Si True, elimina la fila que contiene el marcador <<table_nombre>>.
        Si False, conserva la fila (sobrescribe contenido).
        
        Args:
            eliminar: Boolean para activar/desactivar eliminación
        
        Raises:
            ValueError: Si eliminar no es booleano
        """
        if not isinstance(eliminar, bool):
            raise ValueError(f"eliminar debe ser booleano. Recibido: {eliminar}")
        self._opciones["eliminar_fila_marcador"] = eliminar
    
    # ===== GETTERS =====
    
    def get_aplanar_multiindex(self) -> bool:
        """Retorna si se debe aplanar MultiIndex"""
        return self._opciones.get("aplanar_multiindex", True)
    
    def get_detectar_merge(self) -> bool:
        """Retorna si se debe detectar y aplicar merge"""
        return self._opciones.get("detectar_merge", False)
    
    def get_columnas_para_merge(self) -> Optional[List[int]]:
        """Retorna lista de columnas para merge"""
        return self._opciones.get("columnas_para_merge")
    
    def get_eliminar_fila_marcador(self) -> bool:
        """Retorna si se debe eliminar la fila marcador"""
        return self._opciones.get("eliminar_fila_marcador", True)


class FigSchema:
    """
    Clase para construir un diccionario de figuras con sus atributos.
    
    Atributos:
    ruta: Ruta completa a la imagen (string)
    titulo: Título de la figura (string)
    tamanio: Tamaño de la figura (int, por ejemplo, 6 para 6 cm de ancho)
    bookmark: Bookmark para la figura (string). Debe iniciar con "RefFigura_" seguido del nombre de la figura.
    estilo_figura: Estilo de párrafo (style_id) para el párrafo que contiene la imagen.
    estilo_titulo: Estilo de párrafo (style_id) para el pie de figura.
    
    Ejemplo:
    En la plantilla de word la variable donde se inserta la figura se llamaría <<fig_ejemplo>>, 
    en el excel que tiene los datos necesarios para la plantilla debe existir, en la columna 0, la variable <<fig_ejemplo>>
    y en la misma fila de esa variable, a partir de la columna 1 deben estar los nombres de los archivos de imagen correspondientes a esa figura (sin extensión).
    
    Entonces, este diccionario tendrá las características de la figura. Y deberá usarse para cada una de las figuras asociadas a esa variable. 
    Por ejemplo, si en el excel hay 1 filas con la variable <<fig_ejemplo>>, y esta tiene 3 figuras, cada figura deberá tener la ruta, título, tamaño y bookmark correspondiente a cada imagen. 
    Y el bookmark de cada imagen debe iniciar con "RefFigura_" seguido del nombre del archivo de imagen (sin extensión). 
    Por ejemplo, si el nombre del archivo de imagen es "ejemplo1",
    
    entonces el bookmark para esa figura sería "RefFigura_fig_ejemplo".
    de bookmark: "RefFigura_fig_1"

    Métodos:
    set_ruta: Establece la ruta completa a la imagen a partir de una carpeta, nombre de archivo y extensión.
    set_tamanio: Establece el tamaño de la figura.
    set_bookmark: Establece el bookmark para la figura.
    set_titulo: Establece el título de la figura.
    return_dict: Devuelve un diccionario con los atributos de la figura.
    """
    
    def __init__(
        self,
        ruta_a_figura: str = "",
        titulo: str = "",
        tamanio: int = 6,
        bookmark: str = "",
        estilo_figura: str = "Figura",
        estilo_titulo: str = "Normal",
    ):
        self.ruta = ruta_a_figura
        self.titulo = titulo
        self.tamanio = tamanio
        self.bookmark = bookmark
        self.estilo_figura = estilo_figura
        self.estilo_titulo = estilo_titulo

    def _validar(self):
        """Valida los parámetros de la figura"""
        if not isinstance(self.ruta, str) or not self.ruta:
            raise ValueError("ruta debe ser string no vacío")
        
        if not isinstance(self.tamanio, int) or self.tamanio <= 0:
            raise ValueError("tamanio debe ser número positivo")
        
        # Validar que el archivo existe (opcional, puede comentarse si se prefiere validación lazy)
        if not os.path.exists(self.ruta):
            raise ValueError(f"Archivo de imagen no encontrado: {self.ruta}")
    
        if not self.bookmark.startswith("RefFigura_"):
            # Validar formato de bookmark (opcional)
            raise ValueError("bookmark debe comenzar con 'RefFigura_' para ser válido en referencias cruzadas") 

        if not isinstance(self.estilo_figura, str) or not self.estilo_figura.strip():
            raise ValueError("estilo_figura debe ser string no vacío")

        if not isinstance(self.estilo_titulo, str) or not self.estilo_titulo.strip():
            raise ValueError("estilo_titulo debe ser string no vacío")
    
    def set_ruta(self, ruta_a_la_carpeta_de_imagenes: str, nombre_de_archivo: str, extension: str = "jpg"):
        carpeta = ruta_a_la_carpeta_de_imagenes
        ruta_completa = os.path.join(carpeta, nombre_de_archivo+"."+extension)
        self.ruta = ruta_completa.strip()
    
    def set_tamanio(self, tamanio: int):
        self.tamanio = tamanio
    
    def set_bookmark(self, nombre_de_archivo: str):
        bookmark = "RefFigura_"+nombre_de_archivo
        self.bookmark = bookmark.strip()
        
    def set_titulo(self, titulo_de_figura: str):
        titulo = titulo_de_figura.strip() 
        if titulo != "":
            titulo = titulo if titulo.endswith(".") else titulo + "."
        self.titulo = titulo

    def set_estilo_figura(self, estilo_figura: str):
        self.estilo_figura = estilo_figura.strip()

    def set_estilo_titulo(self, estilo_titulo: str):
        self.estilo_titulo = estilo_titulo.strip()
           
    def to_dict(self) -> dict:
        """
        Retorna diccionario con la configuración de la figura.
        
        Returns:
            Diccionario con ruta, titulo, tamanio, bookmark, estilo_figura y estilo_titulo
        """
        self._validar()
        return {
            "ruta": self.ruta,
            "titulo": self.titulo,
            "tamanio": self.tamanio,
            "bookmark": self.bookmark,
            "estilo_figura": self.estilo_figura,
            "estilo_titulo": self.estilo_titulo,
        }


class TablaSchema:
    """
    Clase para construir el diccionario de configuración de una variable <<nuevatabla_*>>.

    Atributos:
    tabla: DataFrame con los datos de la tabla
    estilos_de_tabla: Diccionario de estilos de tabla
    titulo: Título que se insertará antes de la tabla
    bookmark: Bookmark para referencias cruzadas. Debe iniciar con "RefTabla"

    Métodos requeridos:
    set_tabla, set_titulo, set_bookmark
    """

    def __init__(
        self,
        tabla: Optional[pd.DataFrame] = None,
        estilos_de_tabla: Optional[dict] = None,
        titulo: str = "",
        bookmark: str = "",
    ):
        self.tabla = tabla if tabla is not None else pd.DataFrame()
        self.estilos_de_tabla = estilos_de_tabla if estilos_de_tabla is not None else {}
        self.titulo = titulo
        self.bookmark = bookmark

    def _validar(self):
        if not isinstance(self.tabla, pd.DataFrame):
            raise ValueError("tabla debe ser una instancia de pandas.DataFrame")

        if self.tabla.empty:
            raise ValueError("tabla no puede estar vacía")

        if not isinstance(self.estilos_de_tabla, dict):
            raise ValueError("estilos_de_tabla debe ser un diccionario")

        if not isinstance(self.titulo, str) or not self.titulo.strip():
            raise ValueError("titulo debe ser string no vacío")

        if not isinstance(self.bookmark, str) or not self.bookmark.startswith("RefTabla"):
            raise ValueError("bookmark debe comenzar con 'RefTabla'")

    def set_tabla(self, tabla: pd.DataFrame):
        self.tabla = tabla

    def set_titulo(self, titulo: str):
        self.titulo = titulo.strip()

    def set_bookmark(self, nombre_referencia: str):
        self.bookmark = f"RefTabla_{nombre_referencia}".strip()

    def to_dict(self) -> dict:
        # self._validar()
        return {
            "tabla": self.tabla,
            "estilos_de_tabla": self.estilos_de_tabla,
            "titulo": self.titulo,
            "bookmark": self.bookmark,
        }

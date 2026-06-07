from enum import Enum, unique, auto


@unique
class ColumnsWidthMode(Enum):
    MANUAL = auto() # Para pasar los anchos manualmente

    FROM_LIST = auto()  # Para calcular anchos a partir de una lista simple (e.g., una fila de encabezados)

    FROM_LOL = auto()       # Para calcular anchos a partir de una lista de listas (matriz de datos)
    FROM_LOL_0 = auto() #Para usar los valores del diccionario 0
    FROM_LOL_1 = auto() #Para usar los valores del diccionario 1
    FROM_LOL_2 = auto() #Para usar los valores del diccionario 2
    FROM_LOL_QUANTILE_90 = auto() # Para usar los valores del percentil 90
    FROM_LOL_QUANTILE_90_ONLY_100= auto() # Para usar los valores del percentil 90 
    FROM_LOL_ONLY_100 = auto() # Para usar los 100 primeras listas


    FROM_LOD = auto()   # Para calcular anchos a partir de una lista de diccionarios
    FROM_LOD_0 = auto() #Para usar los valores del diccionario 0
    FROM_LOD_1 = auto() #Para usar los valores del diccionario 1
    FROM_LOD_2 = auto() #Para usar los valores del diccionario 2
    FROM_LOD_KEYS = auto() #Para usar las claves
    FROM_LOD_QUANTILE_90 = auto() # Para usar los valores del percentil 90
    FROM_LOD_QUANTILE_90_ONLY_100 = auto() # Para usar los valores del percentil 90
    FROM_LOD_ONLY_100 = auto() # Para usar los 100 primeros diccionarios

    FROM_SHEET_CELLS= auto()# Como valor se pasa el doc, saca los valores y calcula el width


@unique
class DemoType(Enum):
    SEQUENTIAL = auto()
    CONCURRENT_PROCESS = auto()
    CONCURRENT_THREADS = auto()
    COMMONSERVER_SEQUENTIAL = auto()
    COMMONSERVER_CONCURRENT_PROCESS = auto()
    COMMONSERVER_CONCURRENT_THREADS = auto()

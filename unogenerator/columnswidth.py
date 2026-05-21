from unogenerator import types
from statistics import quantiles



def columnsWidth_from_list(l, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    """
    Calcula el ancho recomendado de las columnas basándose en la longitud máxima
    de los caracteres de una lista simple.

    Retorna una lista de anchos en cm.
    """
    if not l:
        return []

    recommended_widths = []
    for v in l:
        calculated_width = (len(str(v)) * char_to_cm) + padding_cm
        
        # Acotar dentro de los márgenes permitidos
        final_width = max(min_width_cm, min(calculated_width, max_width_cm))
        
        # Redondear para mantener el formato limpio
        recommended_widths.append(round(final_width, 2))

    return recommended_widths


def columnsWidth_from_lol(matrix, n=None, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    """
    Calcula el ancho recomendado de las columnas basándose en la longitud máxima de los caracteres
    de una lista de listas (matriz) dentro de una muestra de 'n' registros.
    
    Toma como máximo los 'n' primeros registros para optimizar el rendimiento.
    Retorna una lista de anchos en cm ordenada por columnas (índice 0, 1, 2...).
    """
    if not matrix or not matrix[0]:
        return []

    sample = matrix if n is None else matrix[:n]


    # 2. Determinar el número de columnas basándonos en la fila más larga del sample
    # (Por si hay filas con longitudes variables)
    num_cols = max(len(row) for row in sample)
    
    # Inicializar una lista de listas para guardar las longitudes de cada columna
    # Ejemplo para 3 columnas: [[], [], []]
    lengths_per_col = [[] for _ in range(num_cols)]

    # 3. Recopilar las longitudes de los caracteres
    for row in sample:
        for col_idx in range(num_cols):
            # Si la fila actual es más corta que num_cols, rellenamos con vacío
            value = row[col_idx] if col_idx < len(row) else ""
            val_str = "" if value is None else str(value)
            lengths_per_col[col_idx].append(len(val_str))

    # 4. Calcular la longitud máxima y convertir a centímetros
    recommended_widths = []
    
    for lengths in lengths_per_col:
        if not lengths:
            max_length = 0
        else:
            max_length = max(lengths)

        # Conversión a centímetros basándonos en el texto
        calculated_width = (max_length * char_to_cm) + padding_cm
        
        # Acotar dentro de los márgenes permitidos
        final_width = max(min_width_cm, min(calculated_width, max_width_cm))
        
        # Redondear para mantener el formato limpio
        recommended_widths.append(round(final_width, 2))

    return recommended_widths


def columnsWidth_from_lod(lod, n=None, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    """
    Calcula el ancho recomendado de las columnas basándose en la longitud máxima
    de los caracteres de una lista de diccionarios (lod), incluyendo la longitud de las claves,
    dentro de una muestra de 'n' registros.
    
    Toma como máximo los 'n' primeros registros para optimizar el rendimiento.
    Retorna una lista de anchos en cm listos para pasar a tu método setColumnsWidth.
    """
    if not lod:
        return []

    sample = lod if n is None else lod[:n]
    # 2. Extraer las claves (columnas) manteniendo el orden del primer diccionario
    keys = list(lod[0].keys())
    
    # Inicializar un diccionario para agrupar las longitudes de cada columna
    # Ejemplo: {'col1': [4, 5, 12, ...], 'col2': [2, 2, 3, ...]}
    lengths_per_col = {key: [] for key in keys}

    # Incluir la longitud de la clave como un posible valor para el ancho de la columna
    for key in keys:
        lengths_per_col[key].append(len(key))

    # 3. Recopilar las longitudes de los caracteres (convertidos a string)
    for row in sample:
        for key in keys:
            # Usamos str() para manejar números, fechas o None de forma segura
            value = row.get(key, "")
            val_str = "" if value is None else str(value)
            lengths_per_col[key].append(len(val_str))
    
    # 4. Calcular la longitud máxima y convertir a centímetros
    recommended_widths = []
    
    for key in keys:
        lengths = lengths_per_col[key]
        
        if not lengths:
            max_length = 0
        else:
            max_length = max(lengths)

        # Convertir caracteres a cm con tus factores de escala
        calculated_width = (max_length * char_to_cm) + padding_cm
        
        # Acotar entre los límites mínimos y máximos
        final_width = max(min_width_cm, min(calculated_width, max_width_cm))
        
        # Redondear a 2 decimales para que quede limpio
        recommended_widths.append(round(final_width, 2))

    return recommended_widths


def columnsWidth_from_lol_with_quantile(matrix, n=100, percentile_value=90, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    """
    Calcula el ancho recomendado de las columnas basándose en un percentil específico
    de la longitud de los caracteres de una lista de listas (matriz) dentro de una muestra de 'n' registros.
    
    Toma como máximo los 'n' primeros registros para optimizar el rendimiento.
    Retorna una lista de anchos en cm ordenada por columnas (índice 0, 1, 2...).
    """
    if not matrix or not matrix[0]:
        return []

    sample = matrix if n is None else matrix[:n]
    num_cols = max(len(row) for row in sample)
    lengths_per_col = [[] for _ in range(num_cols)]

    for row in sample:
        for col_idx in range(num_cols):
            value = row[col_idx] if col_idx < len(row) else ""
            val_str = "" if value is None else str(value)
            lengths_per_col[col_idx].append(len(val_str))

    recommended_widths = []
    
    for lengths in lengths_per_col:
        if not lengths:
            p_length = 0
        elif len(lengths) < 2:
            p_length = lengths[0]
        else:
            # Calculate the specified percentile
            p_length = quantiles(lengths, n=100, method='inclusive')[percentile_value - 1]

        calculated_width = (p_length * char_to_cm) + padding_cm
        final_width = max(min_width_cm, min(calculated_width, max_width_cm))
        recommended_widths.append(round(final_width, 2))

    return recommended_widths


def columnsWidth_from_lod_keys(lod, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    """
    Calcula el ancho recomendado de las columnas basándose únicamente en la longitud
    de las claves del primer diccionario de una lista de diccionarios (lod).
    """
    if not lod:
        return []
    keys = list(lod[0].keys())
    return columnsWidth_from_list(keys, char_to_cm, padding_cm, min_width_cm, max_width_cm)


def columnsWidth_from_lod_with_quantile(lod, n=100, percentile_value=90, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    """
    Calcula el ancho recomendado de las columnas basándose en un percentil específico
    de la longitud de los caracteres de una lista de diccionarios (lod), incluyendo la longitud de las claves,
    dentro de una muestra de 'n' registros.
    
    Toma como máximo los 'n' primeros registros para optimizar el rendimiento.
    Retorna una lista de anchos en cm.
    """
    if not lod:
        return []

    sample = lod[:n]
    keys = list(lod[0].keys())
    lengths_per_col = {key: [len(key)] for key in keys} # Initialize with key lengths

    for row in sample:
        for key in keys:
            value = row.get(key, "")
            val_str = "" if value is None else str(value)
            lengths_per_col[key].append(len(val_str))

    recommended_widths = []
    for key in keys:
        lengths = lengths_per_col[key]
        if not lengths:
            p_length = 0
        elif len(lengths) < 2:
            p_length = lengths[0]
        else:
            p_length = quantiles(lengths, n=100, method='inclusive')[percentile_value - 1]

        calculated_width = (p_length * char_to_cm) + padding_cm
        final_width = max(min_width_cm, min(calculated_width, max_width_cm))
        recommended_widths.append(round(final_width, 2))
    return recommended_widths


def guessColumnsWidth(value: list[dict] | list[list] | list, enummode=types.ColumnsWidthMode.MANUAL, char_to_cm=0.22, padding_cm=0.5, min_width_cm=2.0, max_width_cm=15.0):
    match enummode:
        case types.ColumnsWidthMode.MANUAL:
            return value
        case types.ColumnsWidthMode.FROM_LIST:
            return columnsWidth_from_list(value, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOL:
            return columnsWidth_from_lol(value, None, char_to_cm, padding_cm, min_width_cm, max_width_cm) 
        case types.ColumnsWidthMode.FROM_LOL_0:
            if len(value)==0:
                return []
            return guessColumnsWidth(value[0], types.ColumnsWidthMode.FROM_LIST, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOL_1:
            if len(value)==0:
                return []
            elif len(value)==1:
                return guessColumnsWidth(value, types.ColumnsWidthMode.FROM_LOL_0, char_to_cm, padding_cm, min_width_cm, max_width_cm)
            else:
                return guessColumnsWidth(value[1], types.ColumnsWidthMode.FROM_LIST, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOL_2:
            if len(value)==0:
                return []
            elif len(value)==1:
                return guessColumnsWidth(value, types.ColumnsWidthMode.FROM_LOL_0, char_to_cm, padding_cm, min_width_cm, max_width_cm)
            elif len(value)==2:
                return guessColumnsWidth(value, types.ColumnsWidthMode.FROM_LOL_1, char_to_cm, padding_cm, min_width_cm, max_width_cm)
            else:
                return guessColumnsWidth(value[2], types.ColumnsWidthMode.FROM_LIST, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOL_QUANTILE_90:
            return columnsWidth_from_lol_with_quantile(value, None, 90, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOL_ONLY_100:
            return columnsWidth_from_lol(value, 100, char_to_cm, padding_cm, min_width_cm, max_width_cm) 
        case types.ColumnsWidthMode.FROM_LOL_QUANTILE_90_ONLY_100:
            return columnsWidth_from_lol_with_quantile(value, 100, 90, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        

        case types.ColumnsWidthMode.FROM_LOD:
            return columnsWidth_from_lod(value, None, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_0:
            return columnsWidth_from_list(value[0].values(), char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_1:
            return columnsWidth_from_list(value[1].values(), char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_2:
            return columnsWidth_from_list(value[2].values(), char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_KEYS:
            return columnsWidth_from_lod_keys(value, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_ONLY_100:
            return columnsWidth_from_lod(value, 100, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_QUANTILE_90:
            return columnsWidth_from_lod_with_quantile(value, None, 90, char_to_cm, padding_cm, min_width_cm, max_width_cm)
        case types.ColumnsWidthMode.FROM_LOD_QUANTILE_90_ONLY_100:
            return columnsWidth_from_lod_with_quantile(value, 100, 90, char_to_cm, padding_cm, min_width_cm, max_width_cm)

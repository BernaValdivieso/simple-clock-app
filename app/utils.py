from datetime import datetime, timedelta, time
import pandas as pd

def round_to_nearest(dt, interval):
    """Redondea la hora al intervalo de minutos más cercano."""
    if pd.isna(dt):
        return None
    
    # Convertir a datetime si es solo time
    if isinstance(dt, time):
        dt = datetime.combine(datetime.today(), dt)
    
    minutes = dt.minute
    remainder = minutes % interval
    if remainder < interval / 2:
        rounded_minutes = minutes - remainder
    else:
        rounded_minutes = minutes + (interval - remainder)
    
    # Manejar el caso donde los minutos redondeados exceden 59
    if rounded_minutes >= 60:
        hours_to_add = rounded_minutes // 60
        rounded_minutes = rounded_minutes % 60
        dt = dt + timedelta(hours=hours_to_add)
    
    return dt.replace(minute=rounded_minutes, second=0)

def find_column(df, possible_names):
    """Busca una columna en el DataFrame usando varios nombres posibles."""
    for col in df.columns:
        if any(name.lower() in col.lower() for name in possible_names):
            return col
    return None

def map_column_names(df):
    """Mapea los nombres de columnas del archivo original a los nombres deseados."""
    # Diccionario de mapeo de nombres de columnas
    column_mapping = {
        'Farm name': 'Farm Name',
        'Date': 'Date',
        'Worker ID': 'Worker ID',
        'Job Name': 'Job Name',
        'Job Tags': 'Job Tag',
        'Block name': 'Block Name',
        'Piece Name': 'Piece Name',
        'Clock-In ': 'Rounded Clock-in',
        'Clock-Out': 'Rounded Clock-out',
        '# Hours': '# Hours',
        'Price $/hr': 'Cost per Hour',
        'Price $/piece': 'Price Piece',
        '# Pieces': '# of Pieces',
        'Cost per pieces ($)': 'Cost per Pieces'
    }
    
    # Crear un nuevo DataFrame con las columnas mapeadas
    df_mapped = pd.DataFrame()
    missing_columns = []
    
    # Mapear cada columna del archivo original
    for original_col, target_col in column_mapping.items():
        if original_col in df.columns:
            df_mapped[target_col] = df[original_col]
        else:
            missing_columns.append(original_col)
    
    if missing_columns:
        raise ValueError(f"Las siguientes columnas no se encontraron en el archivo original: {', '.join(missing_columns)}")
    
    return df_mapped

def process_clock_times(df, interval, decimals='all'):
    """Procesa los tiempos de entrada y salida con el intervalo especificado."""
    # Mapear los nombres de columnas
    df_mapped = map_column_names(df)
    
    # Crear una copia del DataFrame con las columnas mapeadas
    df_processed = df_mapped.copy()
    
    # Convertir la columna Date de Excel a formato de fecha legible
    try:
        # Primero intentamos convertir directamente a datetime
        df_processed['Date'] = pd.to_datetime(df_processed['Date'])
    except:
        # Si falla, asumimos que es una fecha de Excel y la convertimos
        df_processed['Date'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(df_processed['Date'].astype(float), unit='D')
    
    df_processed['Date'] = df_processed['Date'].dt.strftime('%m/%d/%Y')
    
    # Convertir las columnas de tiempo a datetime
    df_processed['Rounded Clock-in'] = pd.to_datetime(df_processed['Rounded Clock-in'])
    df_processed['Rounded Clock-out'] = pd.to_datetime(df_processed['Rounded Clock-out'])
    
    # Aplicar redondeo
    df_processed['Rounded Clock-in'] = df_processed['Rounded Clock-in'].apply(
        lambda x: round_to_nearest(x, interval) if pd.notna(x) else None
    )
    df_processed['Rounded Clock-out'] = df_processed['Rounded Clock-out'].apply(
        lambda x: round_to_nearest(x, interval) if pd.notna(x) else None
    )
    
    # Calcular horas trabajadas usando los tiempos redondeados
    def calculate_hours(row):
        if pd.isna(row['Rounded Clock-out']):
            return 0
        
        # Obtener los tiempos redondeados
        clock_in = row['Rounded Clock-in']
        clock_out = row['Rounded Clock-out']
        
        # Si la hora de salida es menor que la de entrada, asumimos que es del día siguiente
        if clock_out.hour < clock_in.hour:
            # Agregar 24 horas a la hora de salida
            hours = (clock_out.hour + 24 - clock_in.hour) + (clock_out.minute - clock_in.minute) / 60
        else:
            hours = (clock_out.hour - clock_in.hour) + (clock_out.minute - clock_in.minute) / 60
        
        # Aplicar el formato de decimales según la selección
        if decimals == 'all':
            return hours  # Retornamos el valor sin redondear
        else:
            return round(hours, int(decimals))
    
    df_processed['# Hours'] = df_processed.apply(calculate_hours, axis=1)
    
    # Convertir las columnas redondeadas a formato de hora sin microsegundos
    df_processed['Rounded Clock-in'] = df_processed['Rounded Clock-in'].apply(
        lambda x: time(x.hour, x.minute, x.second) if pd.notna(x) else None
    )
    df_processed['Rounded Clock-out'] = df_processed['Rounded Clock-out'].apply(
        lambda x: time(x.hour, x.minute, x.second) if pd.notna(x) else None
    )
    
    # Reordenar las columnas según el orden especificado
    final_columns = [
        'Farm Name', 'Date', 'Worker ID', 'Job Name', 'Job Tag',
        'Block Name', 'Piece Name', 'Rounded Clock-in', 'Rounded Clock-out',
        '# Hours', 'Cost per Hour', 'Price Piece', '# of Pieces', 'Cost per Pieces'
    ]
    
    return df_processed[final_columns] 
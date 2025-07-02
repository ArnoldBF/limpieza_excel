import pandas as pd
import json

def cargar_excel(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        return df
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {e}")
        return None

def extraer_json_a_dict(json_str):
    try:
        return json.loads(json_str)
    except Exception:
        return {}

def expandir_columna_json(df, columna_json):
    if columna_json not in df.columns:
        raise ValueError(f"La columna '{columna_json}' no existe en el DataFrame")
    
    json_data = df[columna_json].apply(extraer_json_a_dict)
    json_df = pd.json_normalize(json_data)
    df_expandido = pd.concat([df, json_df], axis=1)
    df_expandido.fillna("", inplace=True)
    return df_expandido

def guardar_excel(df, ruta_salida):
    try:
        df.to_excel(ruta_salida, index=False)
        print(f"Archivo guardado correctamente en '{ruta_salida}'")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

def procesar_excel(ruta_entrada, columna_json, ruta_salida):
    df = cargar_excel(ruta_entrada)
    if df is None:
        return
    try:
        df_expandido = expandir_columna_json(df, columna_json)
        guardar_excel(df_expandido, ruta_salida)
    except Exception as e:
        print(f"Error durante el procesamiento: {e}")

if __name__ == "__main__":
    archivo_entrada = '01.01.2025 a 30.06.2025 data (1).xlsx'
    columna_json = 'custom_attributes'
    archivo_salida = 'archivo_expandido.xlsx'
    
    procesar_excel(archivo_entrada, columna_json, archivo_salida)

import streamlit as st
import pandas as pd
import json
from io import BytesIO

def extraer_json_a_dict(json_str):
    try:
        return json.loads(json_str)
    except:
        return {}

def procesar_excel(file, columna_json):
    df = pd.read_excel(file, engine='openpyxl')
    if columna_json not in df.columns:
        st.error(f"La columna '{columna_json}' no existe en el archivo.")
        return None
    json_data = df[columna_json].apply(extraer_json_a_dict)
    json_df = pd.json_normalize(json_data)
    df_expandido = pd.concat([df, json_df], axis=1)
    df_expandido.fillna("", inplace=True)
    return df_expandido

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

st.title("Expander JSON en Excel")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df_preview = pd.read_excel(uploaded_file, engine='openpyxl', nrows=5)
    st.write("Vista previa del archivo:")
    st.dataframe(df_preview)
    
    columna_json = st.text_input("Nombre de la columna que contiene JSON:", value="custom_attributes")
    
    if st.button("Procesar archivo"):
        with st.spinner("Procesando..."):
            uploaded_file.seek(0)  # volver al inicio del archivo
            df_exp = procesar_excel(uploaded_file, columna_json)
            if df_exp is not None:
                st.success("Archivo procesado correctamente!")
                st.dataframe(df_exp.head())
                
                excel_bytes = to_excel_bytes(df_exp)
                st.download_button(
                    label="Descargar archivo expandido",
                    data=excel_bytes,
                    file_name="archivo_expandido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

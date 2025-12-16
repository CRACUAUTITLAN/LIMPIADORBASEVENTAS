import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Generador Base Power BI", layout="wide")

st.title("🏭 Generador de Base Maestra (Formato Power BI)")
st.markdown("""
**Instrucciones:**
Este programa transforma el reporte de BPro en el formato **EXACTO** que requiere el Dashboard de Ventas.
1. Sube el reporte sucio (.xls).
2. Sube el catálogo de partes (.xlsx).
3. Selecciona el año y descarga.
""")

# --- FUNCIONES DE LIMPIEZA ---

def clean_client_name(text):
    if not isinstance(text, str): return str(text)
    text = text.replace("- -", "").replace("--", "")
    text = re.sub(r'^[\W_]+|[\W_]+$', '', text)
    return text.strip()

def is_valid_date(value):
    if pd.isna(value): return False
    if isinstance(value, (datetime, pd.Timestamp)): return True
    if isinstance(value, str):
        try:
            pd.to_datetime(value, dayfirst=True)
            if '/' in value or '-' in value: return True
        except: return False
    return False

def is_number(s):
    if pd.isna(s): return False
    try:
        float(str(s).replace(',', '')) 
        return True
    except ValueError:
        return False

def safe_get(row, index, default=0):
    try:
        val = row[index]
        if pd.isna(val) or val == "": return default
        return val
    except IndexError: return default

def process_excel_file(uploaded_file):
    # Leer todas las hojas
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    # Memoria Global (Entre hojas)
    current_client = None
    current_fecha = None
    current_tipo = None
    current_almacen = None

    for sheet_name, df in xls.items():
        data = df.values.tolist()
        
        for i, row in enumerate(data):
            col_a = str(safe_get(row, 0, "")).strip()
            col_b_raw = safe_get(row, 1, "")
            col_b_str = str(col_b_raw).strip()
            col_c = safe_get(row, 2, None)
            
            # 1. CLIENTE (Utilidad en Col B)
            if "Utilidad" in col_b_str: 
                current_client = clean_client_name(col_a)
                continue 
            
            # 2. FIN DE BLOQUE
            if col_a.startswith("Totales de:"):
                current_client = None
                continue
            
            # 3. BASURA
            if col_a == "ID Parte" or col_a == "nan" or col_a == "":
                continue

            # 4. FACTURA (Número en Col B)
            if is_number(col_b_raw):
                if is_valid_date(col_c):
                    current_fecha = col_c
                current_tipo = safe_get(row, 3, "")
                current_almacen = safe_get(row, 4, "")
                continue 

            # 5. PRODUCTO
            if current_client: 
                clean_row = {
                    "CLIENTE": current_client,
                    "NP": col_a,             
                    "DESCR": col_b_str,
                    "LINEA": safe_get(row, 2, ""),     
                    "FECHA": current_fecha if current_fecha else "Sin Fecha",
                    "TIPO": current_tipo,    
                    "ALMACEN": current_almacen,   
                    "CANTIDAD": safe_get(row, 3, 0),
                    
                    # --- CORRECCIÓN AQUÍ: Extracción directa ---
                    "COSTO": safe_get(row, 4, 0),   # Columna E
                    "PRECIO": safe_get(row, 5, 0),  # Columna F
                    
                    "COSTO DE VENTA": safe_get(row, 6, 0), # Columna G
                    "VENTA SUBTOTAL": safe_get(row, 7, 0), # Columna H
                    "IVA": safe_get(row, 8, 0),            # Columna I
                    "VENTA TOTAL": safe_get(row, 9, 0),    # Columna J
                    "UTILIDAD": safe_get(row, 10, 0)       # Columna K
                }
                all_clean_rows.append(clean_row)

    return pd.DataFrame(all_clean_rows)

# --- INTERFAZ ---

with st.sidebar:
    st.header("⚙️ Configuración")
    anio_seleccionado = st.radio(
        "¿Qué año quieres descargar?",
        ["Todo (Sin filtro)", "2024", "2025"],
        index=2
    )

col1, col2 = st.columns(2)
with col1:
    file_sucio = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"])
with col2:
    file_actualizador = st.file_uploader('2. Catálogo Partes (.xlsx)', type=["xlsx", "xls"])

if st.button("🚀 GENERAR BASE MAESTRA"):
    if file_sucio and file_actualizador:
        try:
            with st.spinner('Procesando datos y extrayendo costos...'):
                # 1. Limpieza Inicial
                df_limpio = process_excel_file(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron registros. Revisa el archivo de entrada.")
                else:
                    # 2. Cruce con Catálogo
                    df_update = pd.read_excel(file_actualizador)
                    df_limpio['NP'] = df_limpio['NP'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    # Actualizar Línea y llenar Categoría
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                    
                    # 3. Formato de Fechas y Cálculos Temporales
                    df_final['FECHA_DT'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    
                    # --- FILTRO DE AÑO ---
                    if anio_seleccionado == "2024":
                        df_final = df_final[df_final['FECHA_DT'].dt.year == 2024]
                    elif anio_seleccionado == "2025":
                        df_final = df_final[df_final['FECHA_DT'].dt.year == 2025]
                    
                    # --- CÁLCULOS ADICIONALES (Para igualar al Jefe) ---
                    # Generar columnas 'año' y 'mes'
                    df_final['año'] = df_final['FECHA_DT'].dt.year
                    df_final['mes'] = df_final['FECHA_DT'].dt.month
                    
                    # 4. Formato y Ordenamiento Final
                    df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'], ascending=[True, True])
                    df_final['FECHA'] = df_final['FECHA_DT'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                    
                    # LISTA DE ORO (El orden exacto que pediste)
                    cols_finales = [
                        "CLIENTE", "NP", "DESCR", "CATEGORIA", "LINEA", "FECHA", 
                        "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", 
                        "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", 
                        "IVA", "VENTA TOTAL", "UTILIDAD"
                    ]
                    
                    # Verificar que existan todas
                    for col in cols_finales:
                        if col not in df_final.columns:
                            df_final[col] = 0 # Relleno de seguridad
                            
                    df_final = df_final[cols_finales]
                    
                    st.success(f"¡Listo! {len(df_final)} registros procesados. Formato 100% compatible con Power BI.")
                    st.dataframe(df_final.head())
                    
                    # Descarga
                    nombre_archivo = f"Base_Ventas_{anio_seleccionado if anio_seleccionado != 'Todo (Sin filtro)' else 'Completa'}.xlsx"
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Datos')
                    
                    st.download_button(
                        f"📥 Descargar {nombre_archivo}", 
                        data=buffer.getvalue(), 
                        file_name=nombre_archivo, 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error Técnico: {e}")
    else:
        st.warning("⚠️ Por favor sube ambos archivos para continuar.")

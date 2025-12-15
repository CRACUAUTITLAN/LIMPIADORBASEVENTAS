import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Limpieza Multi-Hoja Corregida", layout="wide")

st.title("🛠️ Limpieza de Base de Datos (Fix Continuidad de Hojas)")
st.markdown("""
**Corrección Crítica:**
Se ha habilitado la **Memoria Global**. Ahora, si un cliente empieza en la Hoja 1 y sus productos continúan en la Hoja 2, el programa recordará al cliente y no borrará los datos.
""")

# --- FUNCIONES DE AYUDA ---

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
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    # --- CORRECCIÓN: VARIABLES GLOBALES (MEMORIA PERSISTENTE) ---
    # Las declaramos FUERA del bucle de hojas para que la información 
    # viaje de la Hoja 1 a la Hoja 2, Hoja 3, etc.
    current_client = None
    current_fecha = None
    current_tipo = None
    current_almacen = None

    # Iterar sobre todas las hojas
    for sheet_name, df in xls.items():
        data = df.values.tolist()
        
        # NOTA: Ya no reseteamos las variables aquí. 
        # Si la Hoja 2 es continuación de la Hoja 1, usará el último cliente visto.
        
        for i, row in enumerate(data):
            col_a = str(safe_get(row, 0, "")).strip()
            col_b_raw = safe_get(row, 1, "")
            col_b_str = str(col_b_raw).strip()
            col_c = safe_get(row, 2, None)
            
            # 1. CLIENTE
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

            # 4. FACTURA
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
                    "PART ID": col_a,
                    "DESCR": col_b_str,
                    "LINEA": safe_get(row, 2, ""),     
                    "FECHA": current_fecha if current_fecha else "Sin Fecha",
                    "TIPO_FACTURA": current_tipo, 
                    "ALMACEN": current_almacen,   
                    "CANTIDAD": safe_get(row, 3, 0),  
                    "COSTO DE VENTA": safe_get(row, 6, 0),
                    "VENTA SUBTOTAL": safe_get(row, 7, 0),
                    "IVA": safe_get(row, 8, 0),
                    "VENTA TOTAL": safe_get(row, 9, 0),
                    "UTILIDAD": safe_get(row, 10, 0)
                }
                all_clean_rows.append(clean_row)

    return pd.DataFrame(all_clean_rows)

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    file_sucio = st.file_uploader("1. Sube el Reporte (.xls)", type=["xls"])
with col2:
    file_actualizador = st.file_uploader("2. Sube Categorías (.xlsx)", type=["xlsx", "xls"])

if st.button("PROCESAR CORREGIDO"):
    if file_sucio and file_actualizador:
        try:
            with st.spinner('Procesando con memoria entre hojas...'):
                df_limpio = process_excel_file(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron registros.")
                else:
                    st.success(f"¡Éxito! Registros recuperados: {len(df_limpio)}")
                    
                    df_update = pd.read_excel(file_actualizador)
                    df_limpio['PART ID'] = df_limpio['PART ID'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      left_on='PART ID', right_on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                    
                    df_final['FECHA_ORDEN'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    df_final = df_final.sort_values(by=['FECHA_ORDEN', 'CLIENTE'], ascending=[True, True])
                    df_final['FECHA'] = df_final['FECHA_ORDEN'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                    
                    cols = ["CLIENTE", "PART ID", "DESCR", "LINEA", "CATEGORIA", "FECHA", 
                            "TIPO_FACTURA", "ALMACEN", "CANTIDAD", "COSTO DE VENTA", 
                            "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                    df_final = df_final[cols]
                    
                    st.dataframe(df_final.head())
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Base Completa')
                    
                    st.download_button("📥 Descargar Base Completa", data=buffer.getvalue(), 
                                     file_name="Reporte_Final_Continuo.xlsx", 
                                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Sube ambos archivos.")

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Limpieza Inventario - Modo Seguro", layout="wide")

st.title("🛠️ Limpieza de Base de Datos (Modo Seguro)")
st.markdown("""
**Estado del Sistema:**
1. **Lógica Cliente:** Detecta "Utilidad" en Columna B -> Asigna Cliente de Columna A.
2. **Protección de Datos:** Ya NO elimina productos si falta la fecha. Los guarda siempre.
3. **Orden:** Cronológico y Alfabético.
""")

# --- FUNCIONES DE AYUDA ---

def clean_client_name(text):
    """Limpia el nombre quitando guiones y símbolos extraños"""
    if not isinstance(text, str): return str(text)
    # Quitar guiones comunes
    text = text.replace("- -", "").replace("--", "")
    # Quitar símbolos al inicio/final usando Regex
    text = re.sub(r'^[\W_]+|[\W_]+$', '', text)
    return text.strip()

def is_valid_date(value):
    """Verifica si es fecha válida"""
    if pd.isna(value): return False
    if isinstance(value, (datetime, pd.Timestamp)): return True
    if isinstance(value, str):
        try:
            pd.to_datetime(value, dayfirst=True)
            if '/' in value or '-' in value: return True
        except: return False
    return False

def is_number(s):
    """Detecta si el valor en Col B es numérico (Indica que es una Factura)"""
    if pd.isna(s): return False
    try:
        # Intentamos convertir a float, quitando comas por si es '29,828'
        float(str(s).replace(',', '')) 
        return True
    except ValueError:
        return False

def safe_get(row, index, default=0):
    """Evita que el programa falle si la fila de Excel es muy corta"""
    try:
        val = row[index]
        if pd.isna(val) or val == "":
            return default
        return val
    except IndexError:
        return default

def process_excel_file(uploaded_file):
    # Usamos xlrd para leer archivos .xls antiguos
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    # Iterar sobre todas las hojas
    for sheet_name, df in xls.items():
        data = df.values.tolist()
        
        # Variables de Memoria
        current_client = None
        current_fecha = None
        current_tipo = None
        current_almacen = None
        
        for i, row in enumerate(data):
            # Obtener datos de forma segura (sin error si la fila está vacía)
            col_a = str(safe_get(row, 0, "")).strip()
            col_b_raw = safe_get(row, 1, "")
            col_b_str = str(col_b_raw).strip()
            col_c = safe_get(row, 2, None)
            
            # --- 1. DETECTOR DE CLIENTE (La regla de "Utilidad") ---
            if "Utilidad" in col_b_str: 
                current_client = clean_client_name(col_a)
                continue # Pasamos a la siguiente línea
            
            # --- 2. DETECTOR DE FIN DE BLOQUE ---
            if col_a.startswith("Totales de:"):
                current_client = None
                continue
            
            # --- 3. IGNORAR BASURA ---
            if col_a == "ID Parte" or col_a == "nan" or col_a == "":
                continue

            # --- 4. DETECTOR DE FACTURA (Col B es Número) ---
            # Si B es número (ej: 29,828), es factura.
            if is_number(col_b_raw):
                # Intentamos capturar la fecha si existe
                if is_valid_date(col_c):
                    current_fecha = col_c
                # Si no hay fecha, mantenemos la anterior o dejamos la que está (No borramos nada)
                
                current_tipo = safe_get(row, 3, "")
                current_almacen = safe_get(row, 4, "")
                continue 

            # --- 5. DETECTOR DE PRODUCTO (Col B NO es Número) ---
            # Si llegamos aquí y tenemos un cliente activo, ES UN PRODUCTO.
            if current_client: 
                
                # REGLA BLINDADA: Guardamos SIEMPRE, aunque falten datos como utilidad o fecha.
                clean_row = {
                    "CLIENTE": current_client,
                    "PART ID": col_a,
                    "DESCR": col_b_str,
                    "LINEA": safe_get(row, 2, ""),     
                    "FECHA": current_fecha if current_fecha else "Sin Fecha", # No borramos si falta fecha
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

if st.button("PROCESAR DATOS"):
    if file_sucio and file_actualizador:
        try:
            with st.spinner('Procesando datos sin perder registros...'):
                # 1. Limpieza
                df_limpio = process_excel_file(file_sucio)
                
                if df_limpio.empty:
                    st.error("Alerta: No se encontraron registros. Verifica el archivo.")
                else:
                    st.success(f"¡Éxito! Se recuperaron {len(df_limpio)} productos.")
                    
                    # 2. Carga Actualizador
                    df_update = pd.read_excel(file_actualizador)
                    df_limpio['PART ID'] = df_limpio['PART ID'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    # 3. Cruce
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      left_on='PART ID', right_on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    # 4. Actualización
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                    
                    # 5. Ordenamiento
                    # Convertimos a fecha real solo para ordenar
                    df_final['FECHA_ORDEN'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    df_final = df_final.sort_values(by=['FECHA_ORDEN', 'CLIENTE'], ascending=[True, True])
                    
                    # Aseguramos formato de texto para el Excel final
                    # Si es NaT (Sin Fecha), lo dejamos vacío o como string
                    df_final['FECHA'] = df_final['FECHA_ORDEN'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                    
                    # 6. Selección Final
                    cols = ["CLIENTE", "PART ID", "DESCR", "LINEA", "CATEGORIA", "FECHA", 
                            "TIPO_FACTURA", "ALMACEN", "CANTIDAD", "COSTO DE VENTA", 
                            "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                    df_final = df_final[cols]
                    
                    st.dataframe(df_final.head())
                    
                    # 7. Descarga
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Base Maestra')
                    
                    st.download_button("📥 Descargar Base Completa", data=buffer.getvalue(), 
                                     file_name="Reporte_Completo_Final.xlsx", 
                                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error Técnico: {e}")
            st.write("Detalle del error:", e) # Muestra el error en pantalla para debug
    else:
        st.warning("Sube ambos archivos.")

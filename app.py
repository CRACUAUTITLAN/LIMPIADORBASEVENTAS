import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="Limpieza Inventario Final", layout="wide")

st.title("🛠️ Limpieza de Base de Datos (Versión Final)")
st.markdown("""
**Reglas Definitivas:**
1. **Cliente:** Se detecta cuando la **Columna B** dice "Utilidad". El nombre está en la Columna A.
2. **Factura:** Se detecta cuando la **Columna C** es una FECHA válida.
3. **Limpieza:** Se eliminan guiones (- -) y símbolos del nombre del cliente.
4. **Orden:** Cronológico (Fecha) -> Alfabético (Cliente).
""")

# --- FUNCIONES DE AYUDA ---

def clean_client_name(text):
    """
    Limpia el nombre del cliente:
    - Quita '- -'
    - Quita espacios extra
    - Quita símbolos raros al inicio o final
    """
    if not isinstance(text, str):
        return str(text)
    
    # 1. Quitar el patrón comun "- -"
    text = text.replace("- -", "")
    text = text.replace("--", "")
    
    # 2. Quitar caracteres que no sean letras/numeros al inicio (como * o - sueltos)
    # Regex: Eliminar caracteres no alfanuméricos del inicio y final
    text = re.sub(r'^[\W_]+|[\W_]+$', '', text)
    
    return text.strip()

def is_valid_date(value):
    """Verifica si es fecha válida en Columna C"""
    if pd.isna(value): return False
    if isinstance(value, (datetime, pd.Timestamp)): return True
    if isinstance(value, str):
        try:
            pd.to_datetime(value, dayfirst=True)
            if '/' in value or '-' in value: return True
        except: return False
    return False

def process_excel_file(uploaded_file):
    # Leer archivo con xlrd (para .xls)
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    # Iterar sobre todas las hojas
    for sheet_name, df in xls.items():
        data = df.values.tolist()
        
        # Variables de estado (Memoria)
        current_client = None
        current_fecha = None
        current_tipo = None
        current_almacen = None
        
        for i, row in enumerate(data):
            # Saltamos filas muy cortas
            if len(row) < 3: continue
                
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = str(row[1]).strip() if pd.notna(row[1]) else ""
            col_c = row[2]
            
            # --- 1. DETECTOR DE CLIENTE (LA REGLA DE "UTILIDAD") ---
            # Si Col B dice "Utilidad", Col A es el cliente
            if "Utilidad" in col_b: 
                raw_name = col_a
                # Aplicamos limpieza profunda
                current_client = clean_client_name(raw_name)
                continue # Pasamos a la siguiente fila
            
            # --- 2. DETECTOR DE FIN DE BLOQUE (Totales de...) ---
            if col_a.startswith("Totales de:"):
                current_client = None
                continue
            
            # --- 3. DETECTOR DE FACTURA (Col C es Fecha) ---
            if is_valid_date(col_c):
                current_fecha = col_c
                if len(row) > 4:
                    current_tipo = row[3]
                    current_almacen = row[4]
                continue

            # --- 4. DETECTOR DE PRODUCTO ---
            # Si no es header, ni total, y tenemos cliente y fecha activos
            if col_a != "" and col_a != "ID Parte" and col_a != "nan":
                # Filtro extra: Si Col B es numero, suele ser factura mal detectada, la ignoramos aqui
                # Pero la regla de oro es: ¿Tenemos cliente y fecha seteados?
                
                if current_client and current_fecha:
                     # Verificar longitud para evitar errores de índice
                     if len(row) > 10:
                        clean_row = {
                            "CLIENTE": current_client,
                            "PART ID": col_a,
                            "DESCR": col_b,
                            "LINEA": row[2],     
                            "FECHA": current_fecha, 
                            "TIPO_FACTURA": current_tipo, 
                            "ALMACEN": current_almacen,   
                            "CANTIDAD": row[3],  
                            "COSTO DE VENTA": row[6],
                            "VENTA SUBTOTAL": row[7],
                            "IVA": row[8],
                            "VENTA TOTAL": row[9],
                            "UTILIDAD": row[10]
                        }
                        all_clean_rows.append(clean_row)

    return pd.DataFrame(all_clean_rows)

# --- INTERFAZ DE USUARIO ---

col1, col2 = st.columns(2)
with col1:
    file_sucio = st.file_uploader("1. Sube el Reporte (.xls)", type=["xls"])
with col2:
    file_actualizador = st.file_uploader("2. Sube Categorías (.xlsx)", type=["xlsx", "xls"])

if st.button("PROCESAR DATOS"):
    if file_sucio and file_actualizador:
        try:
            with st.spinner('Limpiando base de datos...'):
                # 1. Procesar
                df_limpio = process_excel_file(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron datos. Revisa si la columna B tiene la palabra 'Utilidad' junto al nombre del cliente.")
                else:
                    st.success(f"Registros encontrados: {len(df_limpio)}")
                    
                    # 2. Cargar Actualizador
                    df_update = pd.read_excel(file_actualizador)
                    df_limpio['PART ID'] = df_limpio['PART ID'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    # 3. Cruce
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      left_on='PART ID', right_on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    # 4. Actualizar Linea y Categoria
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                    
                    # 5. ORDENAMIENTO (Fecha -> Cliente)
                    df_final['FECHA_DT'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'], ascending=[True, True])
                    
                    # Formato final de fecha string
                    df_final['FECHA'] = df_final['FECHA_DT'].dt.strftime('%d/%m/%Y')
                    
                    # 6. Columnas Finales
                    cols = ["CLIENTE", "PART ID", "DESCR", "LINEA", "CATEGORIA", "FECHA", 
                            "TIPO_FACTURA", "ALMACEN", "CANTIDAD", "COSTO DE VENTA", 
                            "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                    df_final = df_final[cols]
                    
                    st.dataframe(df_final.head())
                    
                    # 7. Descargar
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Base 2024-2025')
                    
                    st.download_button("📥 Descargar Excel Limpio", data=buffer.getvalue(), 
                                     file_name="Reporte_Limpio_Final.xlsx", 
                                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Faltan archivos.")

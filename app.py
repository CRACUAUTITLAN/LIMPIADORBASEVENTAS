import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Generador Master BPro", layout="wide")

if 'id_limpieza' not in st.session_state:
    st.session_state['id_limpieza'] = 0

def reset_all():
    st.session_state['id_limpieza'] += 1

with st.sidebar:
    st.header("⚙️ Configuración")
    st.button("🗑️ LIMPIAR / NUEVA BASE", on_click=reset_all, type="primary")
    st.markdown("---")
    sucursal = st.selectbox("Selecciona la Sucursal:", ["TULTITLAN", "CUAUTITLAN"])
    anio_texto = st.text_input("Escribe el año a trabajar:", value="2025")

st.title("🏭 Generador de Base Maestra")
st.markdown("""
**Instrucciones:**
1. Sube los archivos.
2. Genera la base.
3. Descarga.
""")

# --- FUNCIONES AUXILIARES ---
def clean_client_name(text):
    if not isinstance(text, str): return str(text)
    cleaned = re.sub(r'[^\w\s]', '', text) 
    return cleaned.strip()

def clean_text_general(text):
    if not isinstance(text, str): return str(text)
    return text.strip()

def is_number(s):
    if pd.isna(s): return False
    try:
        float(str(s).replace(',', '')) 
        return True
    except ValueError:
        return False

def get_mes_espanol(fecha):
    if pd.isna(fecha): return "DESCONOCIDO"
    try:
        meses = {
            1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
            5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
            9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
        }
        return meses.get(fecha.month, "DESCONOCIDO")
    except:
        return "DESCONOCIDO"

def is_valid_date(value):
    if pd.isna(value): return False
    if isinstance(value, (datetime, pd.Timestamp)): return True
    if isinstance(value, str):
        try:
            pd.to_datetime(value, dayfirst=True)
            if '/' in value or '-' in value: return True
        except:
            return False
    return False

def safe_get(row, index, default=None):
    try:
        val = row[index]
        if pd.isna(val) or val == "": return default
        return val
    except IndexError: return default

# --- PROCESAMIENTO DE VENDEDORES (ESCÁNER DUAL: COLUMNA B Y F) ---
def process_vendedores_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)
    
    mapping_data = [] 
    current_vendedor = "SIN VENDEDOR"
    
    for index, row in df.iterrows():
        item_a = row[0]
        text_a = str(item_a).strip()
        
        if text_a.lower() in ["venta", "nan", ""] or pd.isna(item_a):
            continue
            
        # Si tiene dígitos, es una factura
        if any(char.isdigit() for char in text_a):
            
            # --- ESCÁNER 1: POSITIVOS (COLUMNA B - Índice 1) ---
            val_b_clean = 0.0
            try:
                val_b = row[1] if len(row) > 1 else 0
                val_b_clean = float(str(val_b).replace(',', '').replace('$', '').replace(' ', ''))
            except:
                val_b_clean = 0.0
            
            # Si encontramos dinero en la B (que no sea 0), lo guardamos
            if abs(val_b_clean) > 0.01:
                mapping_data.append({
                    "FACTURA_MATCH": text_a,
                    "MONTO_MATCH": round(val_b_clean, 2),
                    "VENDEDOR": current_vendedor
                })

            # --- ESCÁNER 2: NEGATIVOS (COLUMNA F - Índice 5) ---
            val_f_clean = 0.0
            try:
                val_f = row[5] if len(row) > 5 else 0
                val_f_clean = float(str(val_f).replace(',', '').replace('$', '').replace(' ', ''))
            except:
                val_f_clean = 0.0
                
            # Si encontramos dinero en la F (que no sea 0), TAMBIÉN lo guardamos
            if abs(val_f_clean) > 0.01:
                mapping_data.append({
                    "FACTURA_MATCH": text_a,
                    "MONTO_MATCH": round(val_f_clean, 2),
                    "VENDEDOR": current_vendedor
                })
                
        else:
            # Cambio de Vendedor
            current_vendedor = text_a
            
    return pd.DataFrame(mapping_data)

# --- PROCESAMIENTO PRINCIPAL BPRO ---
def process_main_bpro(uploaded_file):
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    current_client = None
    current_factura = None
    current_fecha = None
    current_tipo = None
    current_almacen = None

    for sheet_name, df in xls.items():
        data = df.values.tolist()
        for row in data:
            col_a_raw = safe_get(row, 0, "")
            col_a_str = str(col_a_raw).strip()
            col_b_raw = safe_get(row, 1, "")
            col_b_str = str(col_b_raw).strip()
            col_c_raw = safe_get(row, 2, None)

            if "Utilidad" in col_b_str:
                current_client = clean_client_name(col_a_str)
                current_factura = None 
                current_fecha = None
                continue

            if col_a_str.startswith("Totales de:"):
                current_client = None
                current_factura = None
                continue

            if is_number(col_b_raw) and is_valid_date(col_c_raw):
                current_factura = col_a_str
                current_fecha = col_c_raw
                current_tipo = safe_get(row, 3, "")
                current_almacen = safe_get(row, 4, "")
                continue

            if col_a_str in ["ID Parte", "nan", "", "Diario de Ventas", "Página"]:
                continue
            if "Diario de Ventas" in col_a_str or "Página" in col_a_str:
                continue

            if current_client and current_factura:
                try:
                    # BPRO Columna H (Index 7) es VENTA SUBTOTAL
                    # Esta es la que usaremos para casar (sea + o -)
                    subtotal_val = safe_get(row, 7, 0)
                    subtotal_clean = float(str(subtotal_val).replace(',', ''))
                    
                    utilidad_val = safe_get(row, 10, 0)
                    utilidad_clean = float(str(utilidad_val).replace(',', ''))
                except:
                    subtotal_clean = 0.0
                    utilidad_clean = 0.0

                clean_row = {
                    "SUCURSAL": "PENDIENTE",
                    "CLIENTE": current_client,
                    "FACTURA": current_factura,
                    "NP": col_a_str,              
                    "DESCR": col_b_str,
                    "LINEA": str(safe_get(row, 2, "")),        
                    "FECHA": current_fecha if current_fecha else "Sin Fecha",
                    "TIPO": current_tipo,     
                    "ALMACEN": current_almacen,   
                    "CANTIDAD": safe_get(row, 3, 0),
                    "COSTO": safe_get(row, 4, 0),
                    "PRECIO": safe_get(row, 5, 0),
                    "COSTO DE VENTA": safe_get(row, 6, 0),
                    "VENTA SUBTOTAL": subtotal_clean,
                    "IVA": safe_get(row, 8, 0),
                    "VENTA TOTAL": safe_get(row, 9, 0),
                    "UTILIDAD": utilidad_clean
                }
                all_clean_rows.append(clean_row)

    return pd.DataFrame(all_clean_rows)

# --- ZONA DE CARGA ---
st.subheader("1. Carga de Archivos")
col1, col2, col3 = st.columns(3)

uid = st.session_state['id_limpieza']

with col1:
    file_sucio = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"], key=f"sucio_{uid}")
with col2:
    file_vendedores = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"vend_{uid}")
with col3:
    file_catalogo = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xlsx", "xls"], key=f"cat_{uid}")

# --- BOTÓN DE EJECUCIÓN ---
if st.button("🚀 GENERAR BASE MAESTRA"):
    if file_sucio and file_catalogo and file_vendedores:
        try:
            with st.spinner('Procesando datos y cruzando información...'):
                
                # 1. Obtener datos de vendedores (Escáner Dual)
                # Esta tabla tendrá tanto los positivos (Col B) como negativos (Col F) en una sola lista
                df_vendedores_map = process_vendedores_file(file_vendedores)
                
                # 2. Obtener datos BPro
                df_limpio = process_main_bpro(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron registros.")
                else:
                    df_limpio['SUCURSAL'] = sucursal
                    
                    # --- CRUCE MAESTRO (MATCH EXACTO) ---
                    df_limpio['FACTURA_MATCH'] = df_limpio['FACTURA'].astype(str).str.strip()
                    df_vendedores_map['FACTURA_MATCH'] = df_vendedores_map['FACTURA_MATCH'].astype(str).str.strip()
                    
                    # Redondeamos todo a 2 decimales para que el match sea perfecto
                    df_limpio['MONTO_MATCH'] = df_limpio['VENTA SUBTOTAL'].round(2)
                    
                    # === MERGE ===
                    # Ahora es simple: Factura + Monto.
                    # No importa si vino de Col B o Col F, ya están todos en 'df_vendedores_map'
                    df_merged = pd.merge(
                        df_limpio,
                        df_vendedores_map,
                        on=['FACTURA_MATCH', 'MONTO_MATCH'],
                        how='left'
                    )
                    
                    # Si no encontró, ponemos "Sin Vendedor"
                    df_merged['VENDEDOR'] = df_merged['VENDEDOR'].fillna("Sin Vendedor")
                    df_limpio = df_merged

                    # --- CONTINUA EL PROCESO ---
                    df_update = pd.read_excel(file_catalogo)
                    df_limpio['NP'] = df_limpio['NP'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    mask_sin_catalogo = df_final['CATEGORIA'].isna()
                    df_final.loc[mask_sin_catalogo, 'DESCR'] = 'X DEFINIR'
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('X DEFINIR')
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    
                    # CORRECCIONES FINALES
                    cols_numericas = ['CANTIDAD', 'COSTO DE VENTA', 'VENTA SUBTOTAL', 'UTILIDAD']
                    for col in cols_numericas:
                        df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

                    # Corregir signos en devoluciones
                    mask_error_bpro = (df_final['CANTIDAD'] < 0) & (df_final['COSTO DE VENTA'] > 0)
                    df_final.loc[mask_error_bpro, 'COSTO DE VENTA'] = df_final.loc[mask_error_bpro, 'COSTO DE VENTA'] * -1
                    df_final.loc[mask_error_bpro, 'UTILIDAD'] = df_final.loc[mask_error_bpro, 'VENTA SUBTOTAL'] - df_final.loc[mask_error_bpro, 'COSTO DE VENTA']
                    
                    # Fechas y Exportación
                    df_final['FECHA_DT'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    if anio_texto.isdigit():
                        anio_num = int(anio_texto)
                        df_final = df_final[df_final['FECHA_DT'].dt.year == anio_num]
                    
                    df_final['año'] = df_final['FECHA_DT'].dt.year
                    df_final['mes'] = df_final['FECHA_DT'].apply(get_mes_espanol)
                    df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'])
                    df_final['FECHA'] = df_final['FECHA_DT'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                    
                    cols_finales = ["SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR", "CATEGORIA", "LINEA", "FECHA", "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                    
                    for col in cols_finales:
                        if col not in df_final.columns: df_final[col] = 0 
                    df_final = df_final[cols_finales]
                    
                    st.success(f"¡Éxito! Vendedores asignados escaneando Positivos (Col B) y Negativos (Col F).")
                    st.dataframe(df_final.head(10))
                    
                    nombre_archivo = f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx"
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Datos')
                    
                    st.download_button(f"📥 Descargar {nombre_archivo}", data=buffer.getvalue(), file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Error Técnico: {e}")
            st.write("Detalle del error:", e)
    else:
        st.warning("⚠️ Por favor sube los 3 archivos para continuar.")

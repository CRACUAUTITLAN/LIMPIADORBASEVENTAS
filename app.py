import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Generador Master BPro", layout="wide")

# --- 1. INICIALIZAR EL CONTADOR DE LIMPIEZA ---
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

# --- PROCESAMIENTO DE VENDEDORES (LEE SUBTOTAL Y UTILIDAD) ---
def process_vendedores_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)
    
    mapping_data = [] 
    current_vendedor = "SIN VENDEDOR"
    
    for index, row in df.iterrows():
        item_a = row[0]
        text_a = str(item_a).strip()
        
        if text_a.lower() in ["venta", "nan", ""] or pd.isna(item_a):
            continue
            
        if any(char.isdigit() for char in text_a):
            # 1. SUBTOTAL (Columna F - Índice 5)
            try:
                val_sub = row[5] if len(row) > 5 else 0
                sub_float = float(str(val_sub).replace(',', '').replace('$', ''))
            except:
                sub_float = 0.0

            # 2. UTILIDAD (Columna N - Índice 13)
            # Nota: Si tu archivo de vendedores tiene la utilidad en otra columna, ajusta el índice 13.
            try:
                val_util = row[13] if len(row) > 13 else 0
                util_float = float(str(val_util).replace(',', '').replace('$', ''))
            except:
                util_float = 0.0
            
            mapping_data.append({
                "FACTURA_MATCH": text_a,
                "SUBTOTAL_MATCH": round(sub_float, 2),
                "UTILIDAD_MATCH": round(util_float, 2),
                "VENDEDOR": current_vendedor
            })
        else:
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
                    # SUBTOTAL (Col H - Index 7)
                    subtotal_val = safe_get(row, 7, 0)
                    subtotal_clean = float(str(subtotal_val).replace(',', ''))
                    
                    # UTILIDAD (Col K - Index 10)
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
                
                # 1. Obtener datos de vendedores (Tabla con Factura, Subtotal y Utilidad)
                df_vendedores_map = process_vendedores_file(file_vendedores)
                
                # 2. Obtener datos BPro
                df_limpio = process_main_bpro(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron registros.")
                else:
                    df_limpio['SUCURSAL'] = sucursal
                    
                    # --- PREPARACIÓN DE LLAVES DE CRUCE ---
                    df_limpio['FACTURA_MATCH'] = df_limpio['FACTURA'].astype(str).str.strip()
                    df_vendedores_map['FACTURA_MATCH'] = df_vendedores_map['FACTURA_MATCH'].astype(str).str.strip()
                    
                    # Llave 1: SUBTOTAL
                    df_limpio['SUBTOTAL_MATCH'] = df_limpio['VENTA SUBTOTAL'].round(2)
                    df_vendedores_map['SUBTOTAL_MATCH'] = df_vendedores_map['SUBTOTAL_MATCH'].round(2)
                    
                    # Llave 2: UTILIDAD
                    df_limpio['UTILIDAD_MATCH'] = df_limpio['UTILIDAD'].round(2)
                    df_vendedores_map['UTILIDAD_MATCH'] = df_vendedores_map['UTILIDAD_MATCH'].round(2)

                    # === ESTRATEGIA DE CRUCE EN CASCADA (FALLBACK) ===
                    
                    # INTENTO 1: Cruce por FACTURA + SUBTOTAL
                    df_merge_1 = pd.merge(
                        df_limpio,
                        df_vendedores_map[['FACTURA_MATCH', 'SUBTOTAL_MATCH', 'VENDEDOR']],
                        on=['FACTURA_MATCH', 'SUBTOTAL_MATCH'],
                        how='left'
                    )
                    
                    # INTENTO 2: Cruce por FACTURA + UTILIDAD
                    df_merge_2 = pd.merge(
                        df_limpio,
                        df_vendedores_map[['FACTURA_MATCH', 'UTILIDAD_MATCH', 'VENDEDOR']],
                        on=['FACTURA_MATCH', 'UTILIDAD_MATCH'],
                        how='left',
                        suffixes=('_sub', '_util')
                    )
                    
                    # === LA LÓGICA MAESTRA ===
                    # Si el Intento 1 encontró vendedor (no es nulo), úsalo.
                    # Si el Intento 1 falló (es nulo o era 0), usa lo que encontró el Intento 2.
                    
                    # Combinamos las columnas: Prioridad Vendedor_sub (Subtotal), si falta, usa Vendedor_util (Utilidad)
                    df_limpio['VENDEDOR'] = df_merge_1['VENDEDOR'].combine_first(df_merge_2['VENDEDOR'])
                    
                    # Si aun así no encuentra nada, poner "Sin Vendedor"
                    df_limpio['VENDEDOR'] = df_limpio['VENDEDOR'].fillna("Sin Vendedor")

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
                    
                    st.success(f"¡Éxito! Base generada con Lógica Cascada (Subtotal -> Utilidad).")
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

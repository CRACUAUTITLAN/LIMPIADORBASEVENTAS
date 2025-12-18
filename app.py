import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Generador Master BPro", layout="wide")

st.title("🏭 Generador de Base Maestra (Con Vendedores y Facturas)")
st.markdown("""
**Instrucciones:**
1. Sube el reporte sucio de BPro (.xls).
2. Sube el reporte de Vendedores (Columna A con Nombres y Facturas).
3. Sube el catálogo de partes (.xlsx).
""")

# --- FUNCIONES AUXILIARES ---

def clean_text(text):
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
    meses = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
    }
    return meses.get(fecha.month, "DESCONOCIDO")

def is_valid_date(value):
    if pd.isna(value): return False
    if isinstance(value, (datetime, pd.Timestamp)): return True
    if isinstance(value, str):
        try:
            pd.to_datetime(value, dayfirst=True)
            if '/' in value or '-' in value: return True
        except: return False
    return False

def safe_get(row, index, default=0):
    try:
        val = row[index]
        if pd.isna(val) or val == "": return default
        return val
    except IndexError: return default

# --- PROCESAMIENTO DE VENDEDORES ---
def process_vendedores_file(uploaded_file):
    """
    Crea un diccionario {FACTURA: VENDEDOR}
    Lógica: Si es texto puro -> Es Vendedor. Si tiene números -> Es Factura asociada.
    """
    df = pd.read_excel(uploaded_file, header=None)
    data = df.iloc[:, 0].tolist() # Solo columna A
    
    mapa_factura_vendedor = {}
    current_vendedor = "SIN VENDEDOR"
    
    for item in data:
        text = str(item).strip()
        if text.lower() == "venta" or text == "nan" or text == "":
            continue
            
        # Lógica para distinguir Nombre de Factura
        # Si tiene algún dígito, asumimos que es FACTURA (Ej: FA00019, A0002)
        if any(char.isdigit() for char in text):
            mapa_factura_vendedor[text] = current_vendedor
        else:
            # Si es solo letras, asumimos que es NOMBRE DE VENDEDOR
            current_vendedor = text
            
    return mapa_factura_vendedor

# --- PROCESAMIENTO PRINCIPAL ---
def process_main_bpro(uploaded_file):
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    current_client = None
    current_factura = None
    next_is_factura = False # Bandera para capturar la fila siguiente al cliente
    
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
            
            # 1. Detectar Cliente
            if "Utilidad" in col_b_str: 
                current_client = clean_text(col_a)
                # Según tu lógica: Debajo del cliente viene la factura
                next_is_factura = True 
                continue 
            
            # 2. Detectar Factura (La fila justo después del cliente)
            if next_is_factura:
                current_factura = clean_text(col_a)
                next_is_factura = False # Apagamos la bandera
                continue

            if col_a.startswith("Totales de:"):
                current_client = None
                current_factura = None
                continue
                
            if col_a == "ID Parte" or col_a == "nan" or col_a == "":
                continue
                
            # Captura de metadatos de encabezado de items (Fecha, Tipo, Almacen)
            if is_number(col_b_raw):
                if is_valid_date(col_c):
                    current_fecha = col_c
                current_tipo = safe_get(row, 3, "")
                current_almacen = safe_get(row, 4, "")
                continue 

            # 3. Captura de Items (Filas de datos)
            if current_client and current_factura: 
                clean_row = {
                    "CLIENTE": current_client,
                    "FACTURA": current_factura, # Nueva Columna
                    "NP": col_a,             
                    "DESCR": col_b_str,
                    "LINEA": safe_get(row, 2, ""),     
                    "FECHA": current_fecha if current_fecha else "Sin Fecha",
                    "TIPO": current_tipo,    
                    "ALMACEN": current_almacen,   
                    "CANTIDAD": safe_get(row, 3, 0),
                    "COSTO": safe_get(row, 4, 0),
                    "PRECIO": safe_get(row, 5, 0),
                    "COSTO DE VENTA": safe_get(row, 6, 0),
                    "VENTA SUBTOTAL": safe_get(row, 7, 0),
                    "IVA": safe_get(row, 8, 0),
                    "VENTA TOTAL": safe_get(row, 9, 0),
                    "UTILIDAD": safe_get(row, 10, 0)
                }
                all_clean_rows.append(clean_row)

    return pd.DataFrame(all_clean_rows)

# --- INTERFAZ ---

with st.sidebar:
    st.header("⚙️ Configuración")
    sucursal = st.selectbox("Selecciona la Sucursal:", ["TULTITLAN", "CUAUTITLAN"])
    anio_texto = st.text_input("Escribe el año a trabajar:", value="2025")

st.subheader("1. Carga de Archivos")
col1, col2, col3 = st.columns(3)

with col1:
    file_sucio = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"])
with col2:
    file_vendedores = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"])
with col3:
    file_catalogo = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xlsx", "xls"])

if st.button("🚀 GENERAR BASE MAESTRA"):
    if file_sucio and file_catalogo and file_vendedores:
        try:
            with st.spinner('Procesando datos y cruzando información...'):
                
                # 1. Crear Mapa de Vendedores
                mapa_vendedores = process_vendedores_file(file_vendedores)
                st.info(f"Se identificaron {len(set(mapa_vendedores.values()))} vendedores distintos.")

                # 2. Limpieza Principal
                df_limpio = process_main_bpro(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron registros en el archivo BPro.")
                else:
                    # 3. Cruce con Vendedores (Usando FACTURA como llave)
                    # Aseguramos que sea string y quitamos espacios para el match
                    df_limpio['FACTURA'] = df_limpio['FACTURA'].astype(str).str.strip()
                    df_limpio['VENDEDOR'] = df_limpio['FACTURA'].map(mapa_vendedores).fillna("NO ASIGNADO")

                    # 4. Cruce con Catálogo de Partes
                    df_update = pd.read_excel(file_catalogo)
                    df_limpio['NP'] = df_limpio['NP'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                    
                    # 5. Fechas y Mes en Letra
                    df_final['FECHA_DT'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    
                    if anio_texto.isdigit():
                        anio_num = int(anio_texto)
                        df_final = df_final[df_final['FECHA_DT'].dt.year == anio_num]
                    
                    df_final['año'] = df_final['FECHA_DT'].dt.year
                    # AQUI ESTA EL CAMBIO A NOMBRE DE MES
                    df_final['mes'] = df_final['FECHA_DT'].apply(get_mes_espanol)
                    
                    df_final['SUCURSAL'] = sucursal
                    
                    # 6. Orden Final de Columnas
                    cols_finales = [
                        "SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR", "CATEGORIA", "LINEA", 
                        "FECHA", "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", 
                        "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", 
                        "IVA", "VENTA TOTAL", "UTILIDAD"
                    ]
                    
                    # Rellenar columnas faltantes
                    for col in cols_finales:
                        if col not in df_final.columns:
                            df_final[col] = 0 
                            
                    df_final = df_final[cols_finales]
                    df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'], ascending=[True, True])
                    
                    # Formato fecha string para excel
                    df_final['FECHA'] = df_final['FECHA_DT'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                    
                    st.success(f"¡Éxito! Base generada con {len(df_final)} registros.")
                    st.dataframe(df_final.head())
                    
                    nombre_archivo = f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx"
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
            st.write("Detalle del error:", e)
    else:
        st.warning("⚠️ Por favor sube los 3 archivos para continuar.")

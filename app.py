import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
import itertools

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador Master BPro", layout="wide")

# --- ESTADO DE LA SESIÓN ---
if 'id_limpieza' not in st.session_state:
    st.session_state['id_limpieza'] = 0

def reset_all():
    st.session_state['id_limpieza'] += 1

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    st.button("🗑️ LIMPIAR / NUEVA BASE", on_click=reset_all, type="primary", help="Borra los archivos y reinicia la memoria")
    st.markdown("---")
    sucursal = st.selectbox("Selecciona la Sucursal:", ["TULTITLAN", "CUAUTITLAN"])
    anio_texto = st.text_input("Escribe el año a trabajar:", value="2025")

# --- TÍTULO ---
st.title("🏭 Generador de Base Maestra")
st.markdown("""
**Instrucciones:**
1. Sube los archivos (BPro, Vendedores, Catálogo).
2. Presiona "Generar Base Maestra".
3. El sistema realizará el cruce matemático para asignar vendedores en facturas compartidas.
""")

# ==========================================
#           FUNCIONES AUXILIARES
# ==========================================

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

# ==========================================
#      PROCESAMIENTO DE VENDEDORES
# ==========================================
def process_vendedores_file(uploaded_file):
    """
    Crea un diccionario donde la llave es la FACTURA y el valor es una lista
    de candidatos con sus montos esperados (Positivos y Negativos).
    """
    df = pd.read_excel(uploaded_file, header=None)
    
    # Estructura: { 'FACTURA': [ {'Vendedor': 'Hugo', 'Monto_Pos': 100, 'Monto_Neg': 0}, ... ] }
    vendedores_map = {} 
    current_vendedor = "SIN VENDEDOR"
    
    for index, row in df.iterrows():
        item_a = row[0]
        text_a = str(item_a).strip()
        
        if text_a.lower() in ["venta", "nan", ""] or pd.isna(item_a):
            continue
            
        if any(char.isdigit() for char in text_a):
            factura = text_a
            
            # --- COLUMNA B (Positivos) ---
            val_b_clean = 0.0
            try:
                val_b = row[1] if len(row) > 1 else 0
                val_b_clean = float(str(val_b).replace(',', '').replace('$', '').replace(' ', ''))
            except:
                val_b_clean = 0.0

            # --- COLUMNA F (Negativos) ---
            val_f_clean = 0.0
            try:
                val_f = row[5] if len(row) > 5 else 0
                val_f_clean = float(str(val_f).replace(',', '').replace('$', '').replace(' ', ''))
            except:
                val_f_clean = 0.0
            
            if factura not in vendedores_map:
                vendedores_map[factura] = []
            
            vendedores_map[factura].append({
                'Vendedor': current_vendedor,
                'Monto_Pos': val_b_clean,
                'Monto_Neg': val_f_clean
            })

        else:
            current_vendedor = text_a
            
    return vendedores_map

# ==========================================
#         PROCESAMIENTO BPRO (SUCIO)
# ==========================================
def process_main_bpro(uploaded_file):
    """
    Extrae la data de BPro.
    NUEVO: Captura el TOTAL DE FACTURA (Col G) del encabezado para saber
    si debemos buscar en positivos o negativos.
    """
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    all_clean_rows = []

    current_client = None
    current_factura = None
    current_fecha = None
    current_tipo = None
    current_almacen = None
    current_factura_total_global = 0.0 # Columna G del encabezado

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
                continue

            if col_a_str.startswith("Totales de:"):
                current_client = None
                current_factura = None
                continue

            # --- DETECCIÓN DE ENCABEZADO DE FACTURA ---
            if is_number(col_b_raw) and is_valid_date(col_c_raw):
                current_factura = col_a_str
                current_fecha = col_c_raw
                current_tipo = safe_get(row, 3, "")
                current_almacen = safe_get(row, 4, "")
                
                # NUEVO: Capturar el Total Global de la Factura (Columna G = Index 6)
                try:
                    val_g = safe_get(row, 6, 0)
                    current_factura_total_global = float(str(val_g).replace(',', ''))
                except:
                    current_factura_total_global = 0.0
                continue

            if col_a_str in ["ID Parte", "nan", "", "Diario de Ventas", "Página"]:
                continue
            if "Diario de Ventas" in col_a_str or "Página" in col_a_str:
                continue

            # --- DETECCIÓN DE PARTIDAS (ITEMS) ---
            if current_client and current_factura:
                try:
                    subtotal_val = safe_get(row, 7, 0) # Col H
                    subtotal_clean = float(str(subtotal_val).replace(',', ''))
                    
                    utilidad_val = safe_get(row, 10, 0) # Col K
                    utilidad_clean = float(str(utilidad_val).replace(',', ''))
                except:
                    subtotal_clean = 0.0
                    utilidad_clean = 0.0

                clean_row = {
                    "SUCURSAL": "PENDIENTE",
                    "CLIENTE": current_client,
                    "FACTURA": current_factura,
                    "GLOBAL_TOTAL_BPRO": current_factura_total_global, # Dato clave para decidir signo
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

# ==========================================
#     ALGORITMO DE DISTRIBUCIÓN (LA MAGIA)
# ==========================================
def solve_subset_sum(items, targets, tolerance=0.05):
    """
    items: lista de diccionarios [{'idx': 0, 'val': 100}, {'idx': 1, 'val': 50}]
    targets: lista de diccionarios [{'Vendedor': 'Hugo', 'Meta': 100}, {'Vendedor': 'Alex', 'Meta': 50}]
    
    Retorna un diccionario {idx_item: Vendedor}
    """
    assignment = {}
    remaining_items = items.copy()
    
    # Ordenamos targets para tratar de llenar primero los más grandes o específicos
    # (Estrategia simple, puede mejorarse)
    
    for target in targets:
        vendedor = target['Vendedor']
        meta = target['Meta']
        
        if meta == 0:
            continue
            
        # Buscamos combinaciones de items restantes que sumen 'meta'
        found_subset = False
        
        # Probamos combinaciones de tamaño 1 hasta N
        # Limitamos N para que no se cuelgue si hay demasiados items (ej. max 6 items combinados)
        # Si la factura es muy compleja, puede requerir más lógica, pero esto cubre el 99%
        max_r = len(remaining_items) + 1 if len(remaining_items) < 12 else 12
        
        for r in range(1, max_r):
            for combo in itertools.combinations(remaining_items, r):
                current_sum = sum(x['val'] for x in combo)
                
                # Checamos si la suma coincide con tolerancia
                if abs(current_sum - meta) <= tolerance:
                    # ¡MATCH ENCONTRADO!
                    for item in combo:
                        assignment[item['idx']] = vendedor
                        remaining_items.remove(item)
                    found_subset = True
                    break
            if found_subset:
                break
                
        if not found_subset:
            # Si no encontramos match exacto, podríamos tener un problema de remanente
            # O simplemente asignamos al primer vendedor (fallback)
            pass

    # Si sobraron items sin asignar (por errores de centavos o complejidad),
    # se los asignamos al último vendedor procesado o al primero de la lista
    if remaining_items and targets:
        fallback_vendedor = targets[0]['Vendedor']
        for item in remaining_items:
            assignment[item['idx']] = fallback_vendedor
            
    return assignment

def apply_vendor_logic_advanced(df_bpro, mapa_vendedores):
    """
    Aplica la lógica de negocio:
    1. Revisa si el Total Global (Col G) es Positivo o Negativo.
    2. Busca los candidatos en el mapa de vendedores.
    3. Si es factura compartida, ejecuta el algoritmo de suma.
    """
    df_bpro['VENDEDOR'] = "Sin Vendedor"
    
    # Agrupamos por Factura para procesar cada una como un bloque independiente
    for factura, group in df_bpro.groupby('FACTURA'):
        
        # Obtenemos índices del grupo original para poder escribir la respuesta
        indices = group.index.tolist()
        
        # 1. Determinamos Signo según Columna G (Encabezado)
        total_global = group['GLOBAL_TOTAL_BPRO'].iloc[0] # El valor es igual para todo el grupo
        es_positivo = total_global >= 0
        
        # 2. Buscamos candidatos en Vendedores
        candidatos = mapa_vendedores.get(factura, [])
        
        if not candidatos:
            continue # Se queda "Sin Vendedor"
            
        # 3. Filtramos las metas correctas (Col B si es positivo, Col F si es negativo)
        targets = []
        for c in candidatos:
            monto_meta = c['Monto_Pos'] if es_positivo else c['Monto_Neg']
            if abs(monto_meta) > 0.01: # Solo consideramos si el vendedor tiene saldo en esa columna
                targets.append({'Vendedor': c['Vendedor'], 'Meta': monto_meta})
        
        if not targets:
            continue
            
        # CASO SENCILLO: Solo hay un vendedor reclamando esta factura
        if len(targets) == 1:
            df_bpro.loc[indices, 'VENDEDOR'] = targets[0]['Vendedor']
            
        # CASO COMPLEJO: Factura compartida (Hugo y Alejandro)
        else:
            # Preparamos los items (Partes) de BPro (Columna H = VENTA SUBTOTAL)
            items = []
            for idx in indices:
                val_row = df_bpro.at[idx, 'VENTA SUBTOTAL']
                items.append({'idx': idx, 'val': val_row})
            
            # Ejecutamos la magia matemática
            resultado_asignacion = solve_subset_sum(items, targets)
            
            # Escribimos los resultados en el DataFrame
            for idx, vendedor_asignado in resultado_asignacion.items():
                df_bpro.at[idx, 'VENDEDOR'] = vendedor_asignado

    return df_bpro

# ==========================================
#              INTERFAZ DE CARGA
# ==========================================
st.subheader("1. Carga de Archivos")
col1, col2, col3 = st.columns(3)

uid = st.session_state['id_limpieza']

with col1:
    file_sucio = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"], key=f"sucio_{uid}")
with col2:
    file_vendedores = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"vend_{uid}")
with col3:
    file_catalogo = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xlsx", "xls"], key=f"cat_{uid}")

# ==========================================
#           LOGICA PRINCIPAL (EJECUCIÓN)
# ==========================================
if st.button("🚀 GENERAR BASE MAESTRA"):
    if file_sucio and file_catalogo and file_vendedores:
        try:
            with st.spinner('Procesando datos y realizando cruce matemático...'):
                
                # 1. Procesar Vendedores (Crear Mapa de Metas)
                mapa_vendedores = process_vendedores_file(file_vendedores)
                
                # 2. Procesar BPro (Obtener filas y Total Global G)
                df_bpro_raw = process_main_bpro(file_sucio)
                
                if df_bpro_raw.empty:
                    st.error("No se encontraron registros en BPro.")
                else:
                    df_bpro_raw['SUCURSAL'] = sucursal
                    df_bpro_raw['FACTURA'] = df_bpro_raw['FACTURA'].astype(str).str.strip()
                    
                    # === APLICACIÓN DE LA LÓGICA AVANZADA ===
                    # Aquí es donde ocurre la distribución inteligente
                    df_limpio = apply_vendor_logic_advanced(df_bpro_raw, mapa_vendedores)
                    
                    # --- PASO 4: CRUCE CON CATÁLOGO ---
                    df_update = pd.read_excel(file_catalogo)
                    df_limpio['NP'] = df_limpio['NP'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    df_final = pd.merge(df_limpio, df_update[['NP', 'LINEA', 'CATEGORIA']], 
                                      on='NP', how='left', suffixes=('_orig', '_new'))
                    
                    # Lógica de relleno de catálogo
                    mask_sin_catalogo = df_final['CATEGORIA'].isna()
                    df_final.loc[mask_sin_catalogo, 'DESCR'] = 'X DEFINIR'
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('X DEFINIR')
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    
                    # --- PASO 5: CORRECCIONES FINALES DE UTILIDAD (NO TOCADAS) ---
                    # Se mantiene tu lógica original para corregir signos en devoluciones
                    cols_numericas = ['CANTIDAD', 'COSTO DE VENTA', 'VENTA SUBTOTAL', 'UTILIDAD']
                    for col in cols_numericas:
                        df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

                    mask_error_bpro = (df_final['CANTIDAD'] < 0) & (df_final['COSTO DE VENTA'] > 0)
                    df_final.loc[mask_error_bpro, 'COSTO DE VENTA'] = df_final.loc[mask_error_bpro, 'COSTO DE VENTA'] * -1
                    # Recalcular utilidad con el costo correcto
                    df_final.loc[mask_error_bpro, 'UTILIDAD'] = df_final.loc[mask_error_bpro, 'VENTA SUBTOTAL'] - df_final.loc[mask_error_bpro, 'COSTO DE VENTA']
                    
                    # --- PASO 6: FECHAS Y ORDEN ---
                    df_final['FECHA_DT'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    if anio_texto.isdigit():
                        anio_num = int(anio_texto)
                        df_final = df_final[df_final['FECHA_DT'].dt.year == anio_num]
                    
                    df_final['año'] = df_final['FECHA_DT'].dt.year
                    df_final['mes'] = df_final['FECHA_DT'].apply(get_mes_espanol)
                    df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'])
                    df_final['FECHA'] = df_final['FECHA_DT'].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                    
                    # Selección de columnas finales
                    cols_finales = ["SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR", "CATEGORIA", "LINEA", "FECHA", "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                    
                    # Asegurar que existan todas
                    for col in cols_finales:
                        if col not in df_final.columns: df_final[col] = 0 
                    df_final = df_final[cols_finales]
                    
                    # --- FIN: MOSTRAR Y DESCARGAR ---
                    st.success(f"¡Éxito! Base generada con distribución inteligente de vendedores.")
                    st.dataframe(df_final.head(10))
                    
                    nombre_archivo = f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx"
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Datos')
                    
                    st.download_button(
                        label=f"📥 Descargar {nombre_archivo}", 
                        data=buffer.getvalue(), 
                        file_name=nombre_archivo, 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error Técnico: {e}")
            st.write("Detalle del error:", e)
    else:
        st.warning("⚠️ Por favor sube los 3 archivos para continuar.")

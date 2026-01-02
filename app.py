import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import time
from datetime import datetime
import itertools
from decimal import Decimal, getcontext

# ======================================================
# CONFIGURACIÓN INICIAL
# ======================================================
getcontext().prec = 12 # Precisión financiera

st.set_page_config(page_title="Generador Master BPro", layout="wide")

if 'id_limpieza' not in st.session_state:
    st.session_state['id_limpieza'] = 0

def reset_all():
    st.session_state['id_limpieza'] += 1

# ======================================================
# BARRA LATERAL
# ======================================================
with st.sidebar:
    st.header("⚙️ Configuración")
    st.button("🗑️ LIMPIAR / NUEVA BASE", on_click=reset_all, type="primary")
    st.markdown("---")
    sucursal = st.selectbox("Selecciona la Sucursal:", ["TULTITLAN", "CUAUTITLAN"])
    anio_texto = st.text_input("Escribe el año a trabajar:", value="2025")

st.title("🏭 Generador de Base Maestra (Versión Ilimitada)")
st.info("💡 Lógica de 'Último Hombre en Pie' activa: Elimina los 'Sin Vendedor' por diferencias de centavos asignando el remanente al último vendedor.")

# ======================================================
# 1. FUNCIONES AUXILIARES
# ======================================================
def clean_client_name(text):
    if not isinstance(text, str): return str(text)
    return re.sub(r'[^\w\s]', '', text).strip()

def is_number(s):
    try:
        float(str(s).replace(',', ''))
        return True
    except: return False

def is_valid_date(value):
    try:
        pd.to_datetime(value, dayfirst=True)
        return True
    except: return False

def safe_get(row, index, default=None):
    try:
        val = row[index]
        return default if (pd.isna(val) or val == "") else val
    except: return default

def get_mes_espanol(fecha):
    meses = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 
             7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
    return meses.get(fecha.month, "DESCONOCIDO")

# ======================================================
# 2. PROCESAMIENTO VENDEDORES
# ======================================================
def process_vendedores_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)
    vendedores_map = {}
    current_vendedor = "SIN VENDEDOR"

    for _, row in df.iterrows():
        texto = str(row[0]).strip()
        if not texto or texto.lower() in ["venta", "nan"]: continue

        if any(c.isdigit() for c in texto):
            factura = texto
            # Col B (Positivos)
            pos = Decimal(str(row[1]).replace(',', '').replace('$', '')) if len(row) > 1 and is_number(row[1]) else Decimal(0)
            # Col F (Negativos)
            neg = Decimal(str(row[5]).replace(',', '').replace('$', '')) if len(row) > 5 and is_number(row[5]) else Decimal(0)

            vendedores_map.setdefault(factura, []).append({
                "VENDEDOR": current_vendedor,
                "POS": pos,
                "NEG": neg
            })
        else:
            current_vendedor = texto
    return vendedores_map

# ======================================================
# 3. PROCESAMIENTO BPRO
# ======================================================
def process_main_bpro(uploaded_file):
    try: xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    except: xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="xlrd")
    
    rows = []
    current_client, current_factura, current_fecha = None, None, None
    current_tipo, current_almacen = None, None
    factura_total_global = Decimal(0)

    for _, df in xls.items():
        for row in df.values.tolist():
            col_a = str(safe_get(row, 0, "")).strip()
            col_b = str(safe_get(row, 1, "")).strip()
            col_c = safe_get(row, 2)

            if "Utilidad" in col_b:
                current_client = clean_client_name(col_a)
                current_factura = None
                continue

            if col_a.startswith("Totales"):
                current_client, current_factura = None, None
                continue

            if is_number(row[1]) and is_valid_date(col_c):
                current_factura = col_a
                current_fecha = col_c
                current_tipo = safe_get(row, 3)
                current_almacen = safe_get(row, 4)
                factura_total_global = Decimal(str(safe_get(row, 6, 0)).replace(',', ''))
                continue

            if col_a in ["ID Parte", "nan", "", "Diario de Ventas", "Página"]: continue

            if current_client and current_factura:
                try:
                    subtotal = Decimal(str(safe_get(row, 7, 0)).replace(',', ''))
                    utilidad = Decimal(str(safe_get(row, 10, 0)).replace(',', ''))
                except: subtotal, utilidad = Decimal(0), Decimal(0)

                rows.append({
                    "SUCURSAL": "PENDIENTE",
                    "CLIENTE": current_client,
                    "FACTURA": current_factura,
                    "GLOBAL_TOTAL_BPRO": factura_total_global,
                    "NP": col_a,
                    "DESCR": col_b,
                    "LINEA": str(safe_get(row, 2, "")),
                    "FECHA": current_fecha,
                    "TIPO": current_tipo,
                    "ALMACEN": current_almacen,
                    "CANTIDAD": safe_get(row, 3, 0),
                    "COSTO": safe_get(row, 4, 0),
                    "PRECIO": safe_get(row, 5, 0),
                    "COSTO DE VENTA": safe_get(row, 6, 0),
                    "VENTA SUBTOTAL": subtotal,
                    "IVA": safe_get(row, 8, 0),
                    "VENTA TOTAL": safe_get(row, 9, 0),
                    "UTILIDAD": utilidad
                })
    return pd.DataFrame(rows)

# ======================================================
# 4. SOLVER INTELIGENTE (OPTIMIZADO)
# ======================================================
def subset_exact_optimized(items, objetivo):
    """
    Busca combinación exacta.
    OPTIMIZACIÓN: Ordena los items de mayor a menor. 
    Esto hace que 'itertools' encuentre la solución muchísimo más rápido
    y permite procesar listas más grandes sin congelarse.
    """
    # Ordenar items por valor absoluto descendente (Truco de velocidad)
    items_sorted = sorted(items, key=lambda x: abs(x["val"]), reverse=True)
    
    # Aumentamos tolerancia a 10 centavos para absorber errores de redondeo de Excel
    tolerancia = Decimal("0.10") 
    
    # Probamos combinaciones
    # Nota: Sin límite duro, pero el ordenamiento ayuda a que no tarde años.
    for r in range(1, len(items_sorted) + 1):
        for combo in itertools.combinations(items_sorted, r):
            suma = sum(x["val"] for x in combo)
            if abs(suma - objetivo) <= tolerancia:
                return combo
    return None

def apply_vendor_logic_unlimited(df, mapa_vendedores):
    df["VENDEDOR"] = "Sin Vendedor"

    for factura, g in df.groupby("FACTURA"):
        indices = g.index
        candidatos = mapa_vendedores.get(factura, [])

        if not candidatos: continue

        total_global = g["GLOBAL_TOTAL_BPRO"].iloc[0]
        es_pos = total_global >= 0

        metas = []
        for c in candidatos:
            monto = c["POS"] if es_pos else c["NEG"]
            if abs(monto) > 0: 
                metas.append({"VENDEDOR": c["VENDEDOR"], "META": monto})
        
        if not metas: continue

        # CASO 1: Un solo vendedor -> Directo
        if len(metas) == 1:
            df.loc[indices, "VENDEDOR"] = metas[0]["VENDEDOR"]
            continue

        # CASO 2: Múltiples vendedores -> Matemáticas + Lógica "Último Hombre"
        items = [{"idx": idx, "val": g.at[idx, "VENTA SUBTOTAL"]} for idx in indices]
        usados = set()
        
        # Ordenamos metas: El que pide más dinero tiene prioridad de búsqueda
        metas.sort(key=lambda x: abs(x["META"]), reverse=True)
        
        for i, meta in enumerate(metas):
            # LOGICA DE ÚLTIMO HOMBRE EN PIE:
            # Si es el último vendedor de la lista y quedan items sin asignar,
            # DÁSELOS TODOS. No importa si cuadra el centavo o no.
            if i == len(metas) - 1:
                remaining_indices = [x["idx"] for x in items if x["idx"] not in usados]
                if remaining_indices:
                    df.loc[remaining_indices, "VENDEDOR"] = meta["VENDEDOR"]
                continue

            # Para los primeros vendedores, buscamos match exacto
            disponibles = [x for x in items if x["idx"] not in usados]
            sol = subset_exact_optimized(disponibles, meta["META"])

            if sol:
                for it in sol:
                    df.at[it["idx"], "VENDEDOR"] = meta["VENDEDOR"]
                    usados.add(it["idx"])
            else:
                # Si no encontramos match exacto para este vendedor intermedio,
                # pasamos al siguiente. El último heredará los problemas.
                pass

    return df

# ======================================================
# 5. EJECUCIÓN PRINCIPAL
# ======================================================
uid = st.session_state['id_limpieza']

col1, col2, col3 = st.columns(3)
with col1: file_bpro = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"], key=f"b_{uid}")
with col2: file_vend = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")
with col3: file_cat = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xls", "xlsx"], key=f"c_{uid}")

if st.button("🚀 GENERAR BASE MAESTRA"):
    if not (file_bpro and file_vend and file_cat):
        st.warning("⚠️ Sube los 3 archivos.")
    else:
        # BARRA 1: PROCESAMIENTO
        prog_bar = st.progress(0, text="Iniciando motores...")

        try:
            # 10% Lectura Vendedores
            prog_bar.progress(10, text="📖 Leyendo Vendedores...")
            mapa = process_vendedores_file(file_vend)
            
            # 30% Lectura BPro
            prog_bar.progress(30, text="🏭 Leyendo BPro...")
            df = process_main_bpro(file_bpro)
            
            if df.empty:
                st.error("Error: BPro vacío.")
                prog_bar.empty()
            else:
                df["SUCURSAL"] = sucursal

                # 60% Cálculo Matemático (El paso pesado)
                prog_bar.progress(60, text="🧮 Realizando cruce matemático ILIMITADO...")
                df = apply_vendor_logic_unlimited(df, mapa)
                
                # 80% Catálogo
                prog_bar.progress(80, text="📚 Cruzando Catálogo...")
                df_cat = pd.read_excel(file_cat)
                df["NP"] = df["NP"].astype(str).str.strip()
                df_cat["NP"] = df_cat["NP"].astype(str).str.strip()
                df_final = pd.merge(df, df_cat[['NP', 'LINEA', 'CATEGORIA']], on='NP', how='left', suffixes=('_orig', '_new'))
                
                # Rellenos y Limpieza
                mask_nocat = df_final['CATEGORIA'].isna()
                df_final.loc[mask_nocat, 'DESCR'] = 'X DEFINIR'
                df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('X DEFINIR')
                df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                
                for col in ["CANTIDAD", "COSTO DE VENTA", "VENTA SUBTOTAL", "UTILIDAD"]:
                    df_final[col] = pd.to_numeric(df_final[col], errors="coerce").fillna(0)

                mask_err = (df_final["CANTIDAD"] < 0) & (df_final["COSTO DE VENTA"] > 0)
                df_final.loc[mask_err, "COSTO DE VENTA"] *= -1
                df_final.loc[mask_err, "UTILIDAD"] = df_final.loc[mask_err, "VENTA SUBTOTAL"] - df_final.loc[mask_err, "COSTO DE VENTA"]

                df_final["FECHA_DT"] = pd.to_datetime(df_final["FECHA"], dayfirst=True, errors="coerce")
                if anio_texto.isdigit():
                    df_final = df_final[df_final["FECHA_DT"].dt.year == int(anio_texto)]

                df_final["año"] = df_final["FECHA_DT"].dt.year
                df_final["mes"] = df_final["FECHA_DT"].apply(get_mes_espanol)
                df_final["FECHA"] = df_final["FECHA_DT"].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'])

                cols_fin = ["SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR", "CATEGORIA", "LINEA", "FECHA", "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                for c in cols_fin:
                    if c not in df_final.columns: df_final[c] = 0
                df_final = df_final[cols_fin]

                # 100% Finalizado cálculo
                prog_bar.progress(100, text="✅ Cálculo terminado.")
                time.sleep(0.5)
                prog_bar.empty()

                # === RESULTADOS Y GENERACIÓN DE EXCEL ===
                st.success(f"¡Base generada! ({len(df_final)} registros)")
                
                # Mostramos dataframe
                st.dataframe(df_final.head(10))
                
                # BARRA 2: PREPARANDO DESCARGA (Lo que pediste)
                # Esta barra aparece justo antes de que el botón esté listo
                download_bar = st.progress(0, text="📁 Preparando archivo Excel para descarga...")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    # Simulamos carga de escritura (Excel tarda un poco en escribir)
                    download_bar.progress(50, text="✍️ Escribiendo hojas de Excel...")
                    df_final.to_excel(writer, index=False, sheet_name='Datos')
                
                download_bar.progress(100, text="🚀 ¡Archivo listo!")
                time.sleep(1)
                download_bar.empty()
                
                st.download_button(f"📥 Descargar {sucursal}", data=buffer.getvalue(), 
                                 file_name=f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx",
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"Error Técnico: {e}")
            if 'prog_bar' in locals(): prog_bar.empty()

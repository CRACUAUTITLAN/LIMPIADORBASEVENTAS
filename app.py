import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import time
from datetime import datetime
import itertools
from decimal import Decimal, getcontext, ROUND_HALF_UP

# ======================================================
# CONFIGURACIÓN INICIAL (Matemática precisa)
# ======================================================
getcontext().prec = 12

st.set_page_config(page_title="Generador Master BPro", layout="wide")

if 'id_limpieza' not in st.session_state:
    st.session_state['id_limpieza'] = 0

def reset_all():
    st.session_state['id_limpieza'] += 1

# ======================================================
# BARRA LATERAL (Estilo Anterior)
# ======================================================
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
st.info("💡 Lógica Activa: Asignación Matemática Exacta (Solver).")

# ======================================================
# 1. FUNCIONES AUXILIARES (Lógica Nueva/Mejorada)
# ======================================================
def clean_client_name(text):
    if not isinstance(text, str): return str(text)
    return re.sub(r'[^\w\s]', '', text).strip()

def clean_currency_str(val):
    if pd.isna(val) or val == "": return Decimal(0)
    s = str(val).replace('$', '').replace(',', '').strip()
    s = re.sub(r'[^\d.-]', '', s) 
    try: return Decimal(s)
    except: return Decimal(0)

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
    if pd.isna(fecha): return "DESCONOCIDO"
    try:
        meses = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 
                 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
        return meses.get(fecha.month, "DESCONOCIDO")
    except: return "DESCONOCIDO"

def is_number(s):
    try: float(str(s).replace(',', '')); return True
    except: return False

# ======================================================
# 2. PROCESAMIENTO VENDEDORES (Lógica Nueva)
# ======================================================
def find_first_amount_in_row(row):
    """Busca el monto en cualquier columna de la fila"""
    for i in range(1, min(len(row), 15)):
        val = clean_currency_str(safe_get(row, i, 0))
        if abs(val) > Decimal("0.01"):
            return val
    return Decimal(0)

def process_vendedores_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)
    vendedores_map = {}
    current_vendedor = "SIN VENDEDOR"
    
    for _, row in df.iterrows():
        texto = str(safe_get(row, 0, "")).strip()
        if not texto or texto.lower() in ["venta", "nan", "total", "iva"]: continue

        if any(c.isdigit() for c in texto):
            factura = texto
            monto = find_first_amount_in_row(row)
            
            if monto >= 0:
                pos = monto
                neg = Decimal(0)
            else:
                pos = Decimal(0)
                neg = monto

            vendedores_map.setdefault(factura, []).append({
                "VENDEDOR": current_vendedor,
                "POS": pos,
                "NEG": neg
            })
        else:
            current_vendedor = texto
            
    return vendedores_map

# ======================================================
# 3. PROCESAMIENTO BPRO (Lógica Nueva - Necesaria para Solver)
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

            val_b = safe_get(row, 1, "")
            if is_valid_date(col_c) and (col_a.startswith("DA") or is_number(val_b)):
                current_factura = col_a
                current_fecha = col_c
                current_tipo = safe_get(row, 3)
                current_almacen = safe_get(row, 4)
                factura_total_global = clean_currency_str(safe_get(row, 6, 0))
                continue

            if col_a in ["ID Parte", "nan", "", "Diario de Ventas", "Página"]: continue

            if current_client and current_factura:
                subtotal = clean_currency_str(safe_get(row, 7, 0))
                utilidad = clean_currency_str(safe_get(row, 10, 0))

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
# 4. SOLVER MATEMÁTICO (Lógica Nueva)
# ======================================================
def subset_sum_solver(items, objetivo):
    tolerancia = Decimal("0.50") 
    objetivo_norm = objetivo.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
    items_sorted = sorted(items, key=lambda x: abs(x["val"]), reverse=True)
    
    for r in range(1, min(len(items_sorted), 8) + 1):
        for combo in itertools.combinations(items_sorted, r):
            suma = sum(x["val"] for x in combo)
            suma_norm = suma.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            if abs(suma_norm - objetivo_norm) <= tolerancia:
                return combo
    return None

# ======================================================
# 5. ASIGNACIÓN INTELIGENTE (Lógica Nueva)
# ======================================================
def apply_vendor_logic_smart(df, mapa_vendedores):
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
            if abs(monto) > Decimal("0.01"): 
                metas.append({"VENDEDOR": c["VENDEDOR"], "META": monto})
        
        if not metas: continue

        if len(metas) == 1:
            df.loc[indices, "VENDEDOR"] = metas[0]["VENDEDOR"]
            continue

        items_todos = []
        for idx in indices:
            raw_val = g.at[idx, "VENTA SUBTOTAL"]
            val_dec = Decimal(str(raw_val)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            items_todos.append({"idx": idx, "val": val_dec})
            
        usados = set()
        metas.sort(key=lambda x: abs(x["META"]), reverse=False) 
        
        for i in range(len(metas) - 1):
            meta_actual = metas[i]
            objetivo = meta_actual["META"]
            sol = None
            
            items_disponibles = [x for x in items_todos if x["idx"] not in usados]
            
            zona_caliente = items_disponibles[-30:] 
            sol = subset_sum_solver(zona_caliente, objetivo)
            
            if not sol:
                pool_global = [x for x in items_disponibles if abs(x["val"]) <= (abs(objetivo) + Decimal("2.00"))]
                sol = subset_sum_solver(pool_global, objetivo)

            if sol:
                for it in sol:
                    df.at[it["idx"], "VENDEDOR"] = meta_actual["VENDEDOR"]
                    usados.add(it["idx"])
            
        vendedor_mayor = metas[-1]["VENDEDOR"]
        remaining_indices = [x["idx"] for x in items_todos if x["idx"] not in usados]
        if remaining_indices:
            df.loc[remaining_indices, "VENDEDOR"] = vendedor_mayor

    return df

# ======================================================
# 6. EJECUCIÓN (Interfaz Antigua + Motor Nuevo)
# ======================================================
uid = st.session_state['id_limpieza']

st.subheader("1. Carga de Archivos")
col1, col2, col3 = st.columns(3)

# Usamos los nombres de variables de la interfaz que te gusta
with col1: file_sucio = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"], key=f"b_{uid}")
with col2: file_vendedores = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")
with col3: file_catalogo = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xlsx", "xls"], key=f"c_{uid}")

if st.button("🚀 GENERAR BASE MAESTRA"):
    if not (file_sucio and file_vendedores and file_catalogo):
        st.warning("⚠️ Por favor sube los 3 archivos para continuar.")
    else:
        # Usamos st.progress como en la lógica nueva porque el proceso matemático es pesado
        # y necesita feedback visual, pero mantenemos el estilo limpio.
        prog_bar = st.progress(0, text="Iniciando motor de cálculo...")
        
        try:
            # 1. Vendedores
            prog_bar.progress(10, text="📖 Leyendo Vendedores...")
            mapa = process_vendedores_file(file_vendedores)
            
            # 2. BPro
            prog_bar.progress(30, text="🏭 Procesando BPro...")
            df = process_main_bpro(file_sucio)
            
            if df.empty:
                st.error("Error: El archivo BPro parece estar vacío o no tiene el formato correcto.")
            else:
                df["SUCURSAL"] = sucursal
                
                # 3. Asignación (El Solver)
                prog_bar.progress(60, text="🧮 Ejecutando Solver Matemático para asignación...")
                df = apply_vendor_logic_smart(df, mapa)
                
                # 4. Catálogo
                prog_bar.progress(80, text="📚 Cruzando Catálogo...")
                df_cat = pd.read_excel(file_catalogo)
                df["NP"] = df["NP"].astype(str).str.strip()
                df_cat["NP"] = df_cat["NP"].astype(str).str.strip()
                df_final = pd.merge(df, df_cat[['NP', 'LINEA', 'CATEGORIA']], on='NP', how='left', suffixes=('_orig', '_new'))
                
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
                
                # Filtro de Año
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

                prog_bar.progress(100, text="✅ Proceso Completado.")
                time.sleep(0.5)
                prog_bar.empty()
                
                # === RESULTADOS (Formato Solicitado) ===
                st.success(f"¡Base generada con éxito! ({len(df_final)} registros)")
                st.dataframe(df_final.head(10)) 
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Datos')
                
                # RESTAURADO: Formato de nombre exacto que pediste
                nombre_archivo = f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx"
                
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo}", 
                    data=buffer.getvalue(), 
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error Técnico: {e}")

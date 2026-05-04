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

# CONFIGURACIÓN INICIAL

# ======================================================

getcontext().prec = 12



# NOTA: Eliminamos st.set_page_config y el Login porque ahora viven en app.py



# ======================================================

# 1. FUNCIONES AUXILIARES (LÓGICA PURA)

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

    meses = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 

             7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}

    return meses.get(fecha.month, "DESCONOCIDO")



def is_number(s):

    try: float(str(s).replace(',', '')); return True

    except: return False



def find_first_amount_in_row(row):

    for i in range(1, min(len(row), 10)): 

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

        if not texto or len(texto) < 3 or texto.lower() in ["venta", "nan", "total", "iva"]: continue



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



def process_main_bpro(uploaded_file):

    try: xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)

    except: xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="xlrd")

    

    rows = []

    current_client, current_factura, current_fecha = None, None, None

    current_tipo, current_almacen = None, None

    factura_total_global = Decimal(0)



    for _, df in xls.items():

        data = df.values.tolist()

        for row in data:

            col_a = str(safe_get(row, 0, "")).strip()

            col_b = str(safe_get(row, 1, "")).strip()

            

            if "Utilidad" in col_b:

                current_client = clean_client_name(col_a)

                current_factura = None

                continue

            if col_a.startswith("Totales"):

                current_client, current_factura = None, None

                continue



            col_c = safe_get(row, 2)

            val_b = safe_get(row, 1, "")

            

            if (col_a.startswith("DA") or is_number(val_b)) and is_valid_date(col_c):

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



def subset_sum_solver_fast(items, objetivo):

    tolerancia = Decimal("0.50")

    objetivo_norm = objetivo.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)

    

    items_validos = [x for x in items if abs(x["val"]) <= (abs(objetivo_norm) + Decimal("2.00"))]

    items_sorted = sorted(items_validos, key=lambda x: abs(x["val"]), reverse=True)[:25]

    max_depth = min(len(items_sorted), 5)

    

    for r in range(1, max_depth + 1):

        for combo in itertools.combinations(items_sorted, r):

            suma = sum(x["val"] for x in combo)

            suma_norm = suma.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)

            if abs(suma_norm - objetivo_norm) <= tolerancia:

                return combo

    return None



def apply_vendor_logic_fast(df, mapa_vendedores):

    df["VENDEDOR"] = "Sin Vendedor"

    grouped = df.groupby("FACTURA")

    

    for factura, g in grouped:

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

            pool = [x for x in items_todos if x["idx"] not in usados]

            sol = subset_sum_solver_fast(pool, objetivo)



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

# 2. FUNCIÓN PRINCIPAL DE VISTA (AQUÍ ESTÁ LA MAGIA)

# ======================================================

def show():

    st.header("🏭 Generador Base Ventas (BPro)")

    st.markdown("---")



    # Gestión de estado para el botón de limpiar

    if 'id_limpieza' not in st.session_state:

        st.session_state['id_limpieza'] = 0



    def reset_all():

        st.session_state['id_limpieza'] += 1



    # --- INTERFAZ DE CONFIGURACIÓN ---

    st.markdown("### 1. Configuración General")

    col_conf1, col_conf2, col_conf3 = st.columns([1, 1, 1])



    with col_conf1:

        sucursal = st.selectbox("🏢 Sucursal:", ["TULTITLAN", "CUAUTITLAN"])



    with col_conf2:

        anio_actual_sistema = str(datetime.now().year)

        anio_texto = st.text_input("📅 Año de Trabajo:", value=anio_actual_sistema)



    with col_conf3:

        st.write("") 

        st.write("") 

        st.button("🗑️ LIMPIAR / NUEVA BASE", on_click=reset_all, type="primary", use_container_width=True)



    st.info("💡 Modo Turbo: Estrategia de asignación prioritaria por monto menor activada.")

    st.markdown("---")



    # --- CARGA DE ARCHIVOS ---

    uid = st.session_state['id_limpieza']

    st.markdown("### 2. Carga de Archivos")



    col1, col2, col3 = st.columns(3)

    with col1: file_bpro = st.file_uploader("1. BPro (.xls)", type=["xls"], key=f"b_{uid}")

    with col2: file_vend = st.file_uploader("2. Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")

    with col3: file_cat = st.file_uploader("3. Catálogo (.xlsx)", type=["xlsx", "xls"], key=f"c_{uid}")



    # --- BOTÓN DE EJECUCIÓN ---

    if st.button("🚀 GENERAR BASE MAESTRA"):

        if not (file_bpro and file_vend and file_cat):

            st.warning("⚠️ Sube los 3 archivos.")

        else:

            prog_bar = st.progress(0, text="Iniciando...")

            try:

                prog_bar.progress(10, text="📖 Leyendo Vendedores...")

                mapa = process_vendedores_file(file_vend)

                

                prog_bar.progress(30, text="🏭 Leyendo BPro...")

                df = process_main_bpro(file_bpro)

                

                if df.empty:

                    st.error("Error: BPro vacío.")

                else:

                    df["SUCURSAL"] = sucursal

                    

                    prog_bar.progress(60, text="🧮 Asignando Vendedores (Motor Turbo)...")

                    df = apply_vendor_logic_fast(df, mapa)

                    

                    prog_bar.progress(80, text="📚 Cruzando Catálogo...")

                    

                    # ==============================================================================

                    # 🛑 ZONA DE SEGURIDAD BLINDADA (ANTIDUPLICADOS) 🛑

                    # ==============================================================================

                    

                    # 1. CARGA SEGURA DEL CATÁLOGO

                    df_cat = pd.read_excel(file_cat)

                    

                    # Convertimos NP a Texto + Quitamos espacios + Borramos Duplicados en Catálogo

                    df_cat["NP"] = df_cat["NP"].astype(str).str.strip()

                    df_cat = df_cat.drop_duplicates(subset=['NP'], keep='first')

                    

                    # 2. LIMPIEZA DE LA BASE BPRO

                    # Convertimos NP a Texto para que coincida con el catálogo

                    df["NP"] = df["NP"].astype(str).str.strip()

                    

                    # Borramos filas EXACTAMENTE idénticas en BPro (fantasmas o errores de carga)

                    df = df.drop_duplicates()

                    

                    # ==============================================================================

                    

                    # 3. CRUCE SEGURO 1 a 1

                    df_final = pd.merge(df, df_cat[['NP', 'LINEA', 'CATEGORIA']], on='NP', how='left', suffixes=('_orig', '_new'))

                    

                    # POST-PROCESAMIENTO

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



                    # ==============================================================================

                    #  🕵️‍♂️ LÓGICA DE REASIGNACIÓN BAJIO

                    # ==============================================================================

                    prog_bar.progress(95, text="🌵 Verificando asignaciones de BAJIO...")



                    BAJIO_VENDEDORES = [

                        "DANIELA OLGUIN MARTINEZ",

                        "JUAN PABLO RODRIGUEZ LOREDO",

                        "PAULINO SANCHEZ MENDEZ"

                    ]



                    BAJIO_ALMACENES = [

                        "ALM. UTEP SAN LUIS",

                        "BISONTE SLP"

                    ]



                    vendedor_check = df_final["VENDEDOR"].astype(str).str.strip().str.upper()

                    almacen_check = df_final["ALMACEN"].astype(str).str.strip().str.upper()



                    mask_bajio = (vendedor_check.isin(BAJIO_VENDEDORES)) | (almacen_check.isin(BAJIO_ALMACENES))



                    cantidad_bajio = mask_bajio.sum()

                    if cantidad_bajio > 0:

                        df_final.loc[mask_bajio, "SUCURSAL"] = "BAJIO"

                        st.toast(f"✅ Se detectaron {cantidad_bajio} registros de BAJIO y se actualizó su sucursal.", icon="🌵")



                    # ==============================================================================

                    

                    cols_fin = ["SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR", "CATEGORIA", "LINEA", "FECHA", "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]

                    for c in cols_fin:

                        if c not in df_final.columns: df_final[c] = 0

                    df_final = df_final[cols_fin]



                    prog_bar.progress(100, text="✅ Listo.")

                    time.sleep(0.5)

                    prog_bar.empty()

                    

                    st.success(f"¡Base generada con éxito! ({len(df_final)} registros)")

                    st.dataframe(df_final.head(10)) 

                    

                    buffer = io.BytesIO()

                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:

                        df_final.to_excel(writer, index=False, sheet_name='Datos')

                    

                    nombre_archivo = f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx"

                    

                    st.download_button(f"📥 Descargar {nombre_archivo}", data=buffer.getvalue(), 

                                         file_name=nombre_archivo,

                                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            except Exception as e:

                st.error(f"Error: {e}")

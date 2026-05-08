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

# Nombres exactos de los dos vendedores especiales
VENDEDOR_A = "HUGO JORGE CORTES RAMIREZ"
VENDEDOR_B = "ALEJANDRO JIMENEZ HERNANDEZ"

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

# ======================================================
# MOTOR DE ASIGNACIÓN
# ======================================================

def decimal_to_cents(d: Decimal) -> int:
    """Convierte Decimal a centavos enteros. Elimina TODO error de punto flotante."""
    return int((d * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

def subset_sum_exact(items: list, target_cents: int):
    """
    Motor exacto en centavos con profundidad limitada a 12.
    Solo se usa para Hugo y Alejandro.

    items        : lista de {"idx": int, "val_cents": int}
    target_cents : objetivo en centavos (entero puro)
    Retorna      : lista de idx que forman el subset exacto, o None si no encuentra.
    """
    if target_cents == 0:
        return []
    total = sum(x["val_cents"] for x in items)
    if target_cents == total:
        return [x["idx"] for x in items]
    if target_cents > total or target_cents < 0:
        return None

    items_validos = [x for x in items if x["val_cents"] <= target_cents]
    items_sorted  = sorted(items_validos, key=lambda x: x["val_cents"], reverse=True)
    max_depth     = min(len(items_sorted), 12)

    for r in range(1, max_depth + 1):
        for combo in itertools.combinations(items_sorted, r):
            if sum(x["val_cents"] for x in combo) == target_cents:
                return [x["idx"] for x in combo]

    return None


def apply_vendor_logic_fast(df, mapa_vendedores):
    df["VENDEDOR"] = "Sin Vendedor"
    facturas_error = []

    for factura, g in df.groupby("FACTURA"):
        indices    = g.index
        candidatos = mapa_vendedores.get(factura, [])
        if not candidatos:
            continue

        es_pos = g["GLOBAL_TOTAL_BPRO"].iloc[0] >= 0

        metas = [
            {"VENDEDOR": c["VENDEDOR"], "META": c["POS"] if es_pos else c["NEG"]}
            for c in candidatos
            if abs(c["POS"] if es_pos else c["NEG"]) > Decimal("0.01")
        ]
        if not metas:
            continue

        # --- Un solo vendedor: asignación directa ---
        if len(metas) == 1:
            df.loc[indices, "VENDEDOR"] = metas[0]["VENDEDOR"]
            continue

        # --- Construir items en centavos ---
        items = []
        for idx in indices:
            raw = Decimal(str(g.at[idx, "VENTA SUBTOTAL"])).quantize(
                      Decimal("0.01"), rounding=ROUND_HALF_UP)
            items.append({"idx": idx, "val_cents": decimal_to_cents(raw)})

        total_items_cents = sum(x["val_cents"] for x in items)

        # ============================================================
        # CASO ESPECIAL: Hugo + Alejandro — Motor exacto en centavos
        # ============================================================
        nombres_en_factura = {m["VENDEDOR"] for m in metas}
        es_caso_especial   = (VENDEDOR_A in nombres_en_factura and
                              VENDEDOR_B in nombres_en_factura and
                              len(metas) == 2)

        if es_caso_especial:
            meta_a_cents = decimal_to_cents(next(m["META"] for m in metas if m["VENDEDOR"] == VENDEDOR_A))
            meta_b_cents = decimal_to_cents(next(m["META"] for m in metas if m["VENDEDOR"] == VENDEDOR_B))

            diferencia = abs(total_items_cents - meta_a_cents - meta_b_cents)
            if diferencia > 2:
                facturas_error.append({
                    "FACTURA" : factura,
                    "PROBLEMA": "inconsistencia_en_datos",
                    "DETALLE" : (
                        f"Items suman ${total_items_cents/100:.2f} | "
                        f"Metas suman ${(meta_a_cents + meta_b_cents)/100:.2f} | "
                        f"Diferencia: ${diferencia/100:.2f}"
                    )
                })
                ganador = max(metas, key=lambda x: x["META"])["VENDEDOR"]
                df.loc[indices, "VENDEDOR"] = ganador
                continue

            # Intenta desde el lado de Hugo
            solucion = subset_sum_exact(items, meta_a_cents)
            if solucion is not None:
                sol_set = set(solucion)
                for item in items:
                    df.at[item["idx"], "VENDEDOR"] = (
                        VENDEDOR_A if item["idx"] in sol_set else VENDEDOR_B
                    )
                continue

            # Intenta desde el lado de Alejandro
            solucion_b = subset_sum_exact(items, meta_b_cents)
            if solucion_b is not None:
                sol_set_b = set(solucion_b)
                for item in items:
                    df.at[item["idx"], "VENDEDOR"] = (
                        VENDEDOR_B if item["idx"] in sol_set_b else VENDEDOR_A
                    )
                continue

            # Ninguna dirección funcionó — registrar para revisión manual
            facturas_error.append({
                "FACTURA" : factura,
                "PROBLEMA": "sin_particion_exacta",
                "DETALLE" : (
                    f"Sin combinación exacta en profundidad 12. "
                    f"Hugo: ${meta_a_cents/100:.2f} | "
                    f"Alejandro: ${meta_b_cents/100:.2f}"
                )
            })
            ganador = max(metas, key=lambda x: x["META"])["VENDEDOR"]
            df.loc[indices, "VENDEDOR"] = ganador
            continue

        # ============================================================
        # TODOS LOS DEMÁS CASOS — Motor rápido original
        # ============================================================
        usados = set()
        metas.sort(key=lambda x: abs(x["META"]), reverse=False)

        for i in range(len(metas) - 1):
            meta_actual   = metas[i]
            objetivo_norm = meta_actual["META"].quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            tolerancia    = Decimal("0.50")
            pool          = [x for x in items if x["idx"] not in usados]

            items_validos = [x for x in pool
                             if abs(Decimal(str(x["val_cents"])) / 100) <= (abs(objetivo_norm) + Decimal("2.00"))]
            items_sorted  = sorted(items_validos, key=lambda x: abs(x["val_cents"]), reverse=True)[:25]
            max_depth     = min(len(items_sorted), 5)
            sol           = None

            for r in range(1, max_depth + 1):
                for combo in itertools.combinations(items_sorted, r):
                    suma = sum(Decimal(str(x["val_cents"])) / 100 for x in combo)
                    if abs(suma.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP) - objetivo_norm) <= tolerancia:
                        sol = combo
                        break
                if sol:
                    break

            if sol:
                for it in sol:
                    df.at[it["idx"], "VENDEDOR"] = meta_actual["VENDEDOR"]
                    usados.add(it["idx"])

        vendedor_mayor    = metas[-1]["VENDEDOR"]
        indices_restantes = [x["idx"] for x in items if x["idx"] not in usados]
        if indices_restantes:
            df.loc[indices_restantes, "VENDEDOR"] = vendedor_mayor

    return df, facturas_error


# ======================================================
# 2. FUNCIÓN PRINCIPAL DE VISTA
# ======================================================
def show():
    st.header("🏭 Generador Base Ventas (BPro)")
    st.markdown("---")

    if 'id_limpieza' not in st.session_state:
        st.session_state['id_limpieza'] = 0

    def reset_all():
        st.session_state['id_limpieza'] += 1

    # --- INTERFAZ DE CONFIGURACIÓN ---
    st.markdown("### 1. Configuración General")

    col_conf1, col_conf2, col_conf3, col_conf4 = st.columns(4)

    with col_conf1:
        sucursal = st.selectbox("🏢 Sucursal:", ["TULTITLAN", "CUAUTITLAN"])

    with col_conf2:
        lista_meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
                       "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        mes_actual_idx = datetime.now().month - 1
        mes_trabajo = st.selectbox("📆 Mes de Trabajo:", lista_meses, index=mes_actual_idx)

    with col_conf3:
        anio_actual_sistema = str(datetime.now().year)
        anio_texto = st.text_input("📅 Año de Trabajo:", value=anio_actual_sistema)

    with col_conf4:
        st.write("")
        st.write("")
        st.button("🗑️ LIMPIAR / NUEVA BASE", on_click=reset_all, type="primary", use_container_width=True)

    st.info(f"💡 Modo Turbo Activado. Procesando afectaciones exclusivas para el mes de: **{mes_trabajo}**")
    st.markdown("---")

    # --- CARGA DE ARCHIVOS DINÁMICA ---
    uid = st.session_state['id_limpieza']
    st.markdown("### 2. Carga de Archivos")

    file_acumulada = None

    if mes_trabajo == "ENERO":
        col1, col2, col3 = st.columns(3)
        with col1: file_bpro = st.file_uploader("1. BPro (.xls)", type=["xls", "xlsx"], key=f"b_{uid}")
        with col2: file_vend = st.file_uploader("2. Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")
        with col3: file_cat  = st.file_uploader("3. Catálogo (.xlsx)", type=["xlsx", "xls"], key=f"c_{uid}")
    else:
        col0, col1, col2, col3 = st.columns(4)
        with col0: file_acumulada = st.file_uploader("0. Base Acumulada (.xlsx)", type=["xlsx", "xls"], key=f"a_{uid}")
        with col1: file_bpro = st.file_uploader("1. BPro Nuevo (.xls)", type=["xls", "xlsx"], key=f"b_{uid}")
        with col2: file_vend = st.file_uploader("2. Vend. Nuevo (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")
        with col3: file_cat  = st.file_uploader("3. Catálogo (.xlsx)", type=["xlsx", "xls"], key=f"c_{uid}")

    # --- BOTÓN DE EJECUCIÓN ---
    if st.button("🚀 GENERAR BASE MAESTRA"):
        if not (file_bpro and file_vend and file_cat):
            st.warning("⚠️ Sube los archivos requeridos (BPro, Vendedores, Catálogo).")
        elif mes_trabajo != "ENERO" and not file_acumulada:
            st.warning(f"⚠️ Has seleccionado el mes de {mes_trabajo}. Es OBLIGATORIO subir tu 'Base Acumulada' en la primera casilla.")
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

                    prog_bar.progress(60, text="🧮 Asignando Vendedores...")
                    df, errores_vendedor = apply_vendor_logic_fast(df, mapa)

                    if errores_vendedor:
                        df_err = pd.DataFrame(errores_vendedor)
                        st.warning(f"⚠️ {len(errores_vendedor)} factura(s) requieren revisión manual:")
                        st.dataframe(df_err, use_container_width=True)

                    prog_bar.progress(80, text="📚 Cruzando Catálogo...")

                    # 1. CARGA SEGURA DEL CATÁLOGO
                    df_cat = pd.read_excel(file_cat)
                    df_cat["NP"] = df_cat["NP"].astype(str).str.strip()
                    df_cat = df_cat.drop_duplicates(subset=['NP'], keep='first')

                    # 2. LIMPIEZA DE LA BASE BPRO
                    df["NP"] = df["NP"].astype(str).str.strip()
                    df = df.drop_duplicates()

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

                    # ==============================================================================
                    # LÓGICA DE INTEGRACIÓN MENSUAL (BASE ACUMULADA)
                    # ==============================================================================
                    if mes_trabajo != "ENERO":
                        prog_bar.progress(90, text=f"🔄 Integrando con Acumulado (Sobrescribiendo {mes_trabajo})...")
                        df_acum = pd.read_excel(file_acumulada)

                        df_acum = df_acum[df_acum["mes"] != mes_trabajo]
                        df_final = pd.concat([df_acum, df_final], ignore_index=True)
                        df_final["FECHA_DT"] = pd.to_datetime(df_final["FECHA"], dayfirst=True, errors="coerce")

                    # ORDEN FINAL Y LIMPIEZA DE COLUMNAS
                    df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'])

                    cols_fin = ["SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR",
                                "CATEGORIA", "LINEA", "FECHA", "año", "mes", "TIPO", "ALMACEN",
                                "CANTIDAD", "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL",
                                "IVA", "VENTA TOTAL", "UTILIDAD"]
                    for c in cols_fin:
                        if c not in df_final.columns: df_final[c] = 0
                    df_final = df_final[cols_fin]

                    prog_bar.progress(100, text="✅ Listo.")
                    time.sleep(0.5)
                    prog_bar.empty()

                    st.success(f"¡Base de {mes_trabajo} generada e integrada con éxito! (Total: {len(df_final)} registros)")
                    st.dataframe(df_final.tail(10))

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Datos')

                    nombre_archivo = f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx"

                    st.download_button(f"📥 Descargar {nombre_archivo}", data=buffer.getvalue(),
                                         file_name=nombre_archivo,
                                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error procesando la información: {e}")

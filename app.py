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

st.title("🏭 Generador de Base Maestra (Versión Blindada)")
st.info("💡 Lógica Activa: Limpieza agresiva de montos + Búsqueda cronológica inversa.")

# ======================================================
# 1. FUNCIONES AUXILIARES
# ======================================================
def clean_client_name(text):
    if not isinstance(text, str): return str(text)
    return re.sub(r'[^\w\s]', '', text).strip()

def clean_currency_str(val):
    """Limpia strings de moneda agresivamente para evitar errores de lectura."""
    if pd.isna(val) or val == "": return Decimal(0)
    s = str(val).replace('$', '').replace(',', '').strip()
    # Eliminar cualquier caracter que no sea numero, punto o negativo
    s = re.sub(r'[^\d.-]', '', s)
    try:
        return Decimal(s)
    except:
        return Decimal(0)

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
    # Leemos sin header para controlar índices manualmente
    df = pd.read_excel(uploaded_file, header=None)
    vendedores_map = {}
    current_vendedor = "SIN VENDEDOR"

    for _, row in df.iterrows():
        # Tomamos la columna 0 como texto base
        texto = str(safe_get(row, 0, "")).strip()
        
        # Saltamos filas vacías o basura
        if not texto or texto.lower() in ["venta", "nan", "total", "iva"]: continue

        # Detectamos si es Factura (tiene dígitos) o Vendedor (solo letras)
        if any(c.isdigit() for c in texto):
            factura = texto
            
            # LECTURA BLINDADA DE MONTOS
            # Asumimos Col 1 = Venta Positiva, Col 5 = Devolución (Negativa)
            # Usamos clean_currency_str para asegurar que "114.26" se lea bien
            pos_val = safe_get(row, 1, 0)
            neg_val = safe_get(row, 5, 0)
            
            pos = clean_currency_str(pos_val)
            neg = clean_currency_str(neg_val)

            vendedores_map.setdefault(factura, []).append({
                "VENDEDOR": current_vendedor,
                "POS": pos,
                "NEG": neg
            })
        else:
            # Si no tiene números, asumimos que es el encabezado de un Vendedor
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

            # Detectar encabezado de factura
            # Usamos clean_currency_str para el total global también
            val_b = safe_get(row, 1, "")
            if clean_currency_str(val_b) != 0 and is_valid_date(col_c):
                current_factura = col_a
                current_fecha = col_c
                current_tipo = safe_get(row, 3)
                current_almacen = safe_get(row, 4)
                factura_total_global = clean_currency_str(safe_get(row, 6, 0))
                continue

            if col_a in ["ID Parte", "nan", "", "Diario de Ventas", "Página"]: continue

            if current_client and current_factura:
                # Lectura de montos con limpieza
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
                    "VENTA SUBTOTAL": subtotal, # Decimal limpio
                    "IVA": safe_get(row, 8, 0),
                    "VENTA TOTAL": safe_get(row, 9, 0),
                    "UTILIDAD": utilidad
                })
    return pd.DataFrame(rows)

# ======================================================
# 4. SOLVER MATEMÁTICO (SUMA DE SUBCONJUNTOS)
# ======================================================
def subset_sum_solver(items, objetivo):
    """
    items: lista de dicts [{'idx': 10, 'val': Decimal('100.50')}]
    objetivo: Decimal('114.26')
    """
    # Tolerancia ampliada a 50 centavos por temas de redondeo IVA
    tolerancia = Decimal("0.50")
    
    # 1. Normalizar inputs a 2 decimales para comparación estricta
    objetivo_norm = objetivo.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
    
    # Ordenamos descendente para optimizar
    items_sorted = sorted(items, key=lambda x: abs(x["val"]), reverse=True)
    
    # Probamos combinaciones de 1 hasta 6 elementos
    # (Raramente un vendedor captura más de 6 lineas separadas para llegar a un total)
    for r in range(1, min(len(items_sorted), 7) + 1):
        for combo in itertools.combinations(items_sorted, r):
            suma = sum(x["val"] for x in combo)
            suma_norm = suma.quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            
            if abs(suma_norm - objetivo_norm) <= tolerancia:
                return combo
    return None

# ======================================================
# 5. LÓGICA DE ASIGNACIÓN (BÚSQUEDA INVERSA)
# ======================================================
def apply_vendor_logic_smart(df, mapa_vendedores):
    df["VENDEDOR"] = "Sin Vendedor"

    for factura, g in df.groupby("FACTURA"):
        indices = g.index
        candidatos = mapa_vendedores.get(factura, [])

        # Si no hay match en el archivo de vendedores, saltamos
        if not candidatos: continue

        # Determinamos si buscamos POS o NEG basado en el total de la factura
        total_global = g["GLOBAL_TOTAL_BPRO"].iloc[0]
        es_pos = total_global >= 0

        # Preparamos metas
        metas = []
        for c in candidatos:
            monto = c["POS"] if es_pos else c["NEG"]
            # IMPORTANTE: Solo agregamos si el monto es mayor a 0.01
            if abs(monto) > Decimal("0.01"): 
                metas.append({"VENDEDOR": c["VENDEDOR"], "META": monto})
        
        if not metas: continue

        # CASO 1: Vendedor Único -> Asignación directa
        if len(metas) == 1:
            df.loc[indices, "VENDEDOR"] = metas[0]["VENDEDOR"]
            continue

        # CASO 2: Múltiples Vendedores
        
        # Obtenemos items conservando orden original
        # Forzamos conversión a Decimal limpio
        items_todos = []
        for idx in indices:
            raw_val = g.at[idx, "VENTA SUBTOTAL"]
            val_dec = Decimal(str(raw_val)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            items_todos.append({"idx": idx, "val": val_dec})
            
        usados = set()
        
        # Ordenamos metas de MENOR a MAYOR (Alejandro primero)
        metas.sort(key=lambda x: abs(x["META"]), reverse=False)
        
        # Iteramos sobre todos MENOS el último
        for i in range(len(metas) - 1):
            meta_actual = metas[i]
            objetivo = meta_actual["META"]
            sol = None
            
            items_disponibles = [x for x in items_todos if x["idx"] not in usados]
            
            # --- ESTRATEGIA A: ZONA CALIENTE (Últimos 40 items) ---
            # Ampliamos a 40 para no fallar
            zona_caliente = items_disponibles[-40:] 
            sol = subset_sum_solver(zona_caliente, objetivo)
            
            # --- ESTRATEGIA B: BÚSQUEDA GLOBAL ---
            if not sol:
                # Si falla, busca en todo con filtro de monto máximo
                pool_global = [
                    x for x in items_disponibles
                    if abs(x["val"]) <= (abs(objetivo) + Decimal("2.00"))
                ]
                sol = subset_sum_solver(pool_global, objetivo)

            # Aplicar asignación
            if sol:
                for it in sol:
                    df.at[it["idx"], "VENDEDOR"] = meta_actual["VENDEDOR"]
                    usados.add(it["idx"])
            
        # El resto para el Vendedor Mayor (Hugo)
        vendedor_mayor = metas[-1]["VENDEDOR"]
        remaining_indices = [x["idx"] for x in items_todos if x["idx"] not in usados]
        
        if remaining_indices:
            df.loc[remaining_indices, "VENDEDOR"] = vendedor_mayor

    return df

# ======================================================
# 6. EJECUCIÓN PRINCIPAL
# ======================================================
uid = st.session_state['id_limpieza']

col1, col2, col3 = st.columns(3)
with col1: file_bpro = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"], key=f"b_{uid}")
with col2: file_vend = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")
with col3: file_cat = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xlsx", "xls"], key=f"c_{uid}")

if st.button("🚀 GENERAR BASE MAESTRA"):
    if not (file_bpro and file_vend and file_cat):
        st.warning("⚠️ Sube los 3 archivos.")
    else:
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

                # 60% Cálculo Matemático
                prog_bar.progress(60, text="🧮 Asignando Vendedores (Lógica Blindada)...")
                df = apply_vendor_logic_smart(df, mapa)
                
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

                # 100% Finalizado
                prog_bar.progress(100, text="✅ Cálculo terminado.")
                time.sleep(0.5)
                prog_bar.empty()

                # === RESULTADOS ===
                st.success(f"¡Base generada! ({len(df_final)} registros)")
                st.dataframe(df_final.head(10))
                
                download_bar = st.progress(0, text="📁 Preparando descarga...")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    download_bar.progress(50, text="✍️ Escribiendo Excel...")
                    df_final.to_excel(writer, index=False, sheet_name='Datos')
                
                download_bar.progress(100, text="🚀 ¡Listo!")
                time.sleep(1)
                download_bar.empty()
                
                st.download_button(f"📥 Descargar {sucursal}", data=buffer.getvalue(), 
                                 file_name=f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx",
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"Error Técnico: {e}")
            if 'prog_bar' in locals(): prog_bar.empty()

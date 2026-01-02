import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import time  # <--- AGREGA ESTO
from datetime import datetime
import itertools
from decimal import Decimal, getcontext

# ======================================================
# CONFIGURACIÓN INICIAL
# ======================================================
# Configurar precisión para manejo de dinero exacto en Decimals
getcontext().prec = 12

st.set_page_config(page_title="Generador Master BPro", layout="wide")

if 'id_limpieza' not in st.session_state:
    st.session_state['id_limpieza'] = 0

def reset_all():
    st.session_state['id_limpieza'] += 1

# ======================================================
# BARRA LATERAL (SIDEBAR)
# ======================================================
with st.sidebar:
    st.header("⚙️ Configuración")
    st.button("🗑️ LIMPIAR / NUEVA BASE", on_click=reset_all, type="primary")
    st.markdown("---")
    sucursal = st.selectbox("Selecciona la Sucursal:", ["TULTITLAN", "CUAUTITLAN"])
    anio_texto = st.text_input("Escribe el año a trabajar:", value="2025")

st.title("🏭 Generador de Base Maestra (Versión Final)")
st.info("💡 Lógica Híbrida Activa: Cálculo matemático exacto para facturas compartidas + Protección de rendimiento para facturas masivas.")

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
# 2. PROCESAMIENTO DE VENDEDORES (Escáner Dual)
# ======================================================
def process_vendedores_file(uploaded_file):
    """
    Lee el archivo de agentes.
    Detecta montos en Columna B (Positivos/Ventas) y Columna F (Negativos/Devoluciones).
    Crea un mapa maestro {Factura: [Lista de Vendedores y sus Metas]}
    """
    df = pd.read_excel(uploaded_file, header=None)
    vendedores_map = {}
    current_vendedor = "SIN VENDEDOR"

    for _, row in df.iterrows():
        texto = str(row[0]).strip()
        
        # Saltar basura
        if not texto or texto.lower() in ["venta", "nan"]: 
            continue

        # Si tiene dígitos, asumimos que es FACTURA
        if any(c.isdigit() for c in texto):
            factura = texto
            
            # Leer Columna B (Index 1) -> Ventas Positivas
            pos = Decimal(str(row[1]).replace(',', '').replace('$', '')) if len(row) > 1 and is_number(row[1]) else Decimal(0)
            
            # Leer Columna F (Index 5) -> Devoluciones/Negativos
            neg = Decimal(str(row[5]).replace(',', '').replace('$', '')) if len(row) > 5 and is_number(row[5]) else Decimal(0)

            vendedores_map.setdefault(factura, []).append({
                "VENDEDOR": current_vendedor,
                "POS": pos,
                "NEG": neg
            })
        else:
            # Si es texto sin números, es el NOMBRE DEL VENDEDOR
            current_vendedor = texto

    return vendedores_map

# ======================================================
# 3. PROCESAMIENTO BPRO (Lectura Raw)
# ======================================================
def process_main_bpro(uploaded_file):
    """
    Lee BPro. Captura el TOTAL GLOBAL (Col G) para saber el signo de la factura.
    Captura VENTA SUBTOTAL (Col H) de cada partida para hacer el match matemático.
    """
    try: 
        xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    except: 
        xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="xlrd")
    
    rows = []
    current_client, current_factura, current_fecha = None, None, None
    current_tipo, current_almacen = None, None
    factura_total_global = Decimal(0)

    for _, df in xls.items():
        for row in df.values.tolist():
            col_a = str(safe_get(row, 0, "")).strip()
            col_b = str(safe_get(row, 1, "")).strip()
            col_c = safe_get(row, 2)

            # Detectar Cliente
            if "Utilidad" in col_b:
                current_client = clean_client_name(col_a)
                current_factura = None
                continue

            if col_a.startswith("Totales"):
                current_client, current_factura = None, None
                continue

            # Detectar Encabezado Factura (Tiene Total Global en Col G / Index 6)
            if is_number(row[1]) and is_valid_date(col_c):
                current_factura = col_a
                current_fecha = col_c
                current_tipo = safe_get(row, 3)
                current_almacen = safe_get(row, 4)
                
                # CLAVE: Capturamos el total de la factura para saber si es + o -
                factura_total_global = Decimal(str(safe_get(row, 6, 0)).replace(',', ''))
                continue

            if col_a in ["ID Parte", "nan", "", "Diario de Ventas", "Página"]: continue

            # Detectar Partidas (Items)
            if current_client and current_factura:
                try:
                    subtotal = Decimal(str(safe_get(row, 7, 0)).replace(',', '')) # Col H
                    utilidad = Decimal(str(safe_get(row, 10, 0)).replace(',', '')) # Col K
                except: 
                    subtotal, utilidad = Decimal(0), Decimal(0)

                rows.append({
                    "SUCURSAL": "PENDIENTE",
                    "CLIENTE": current_client,
                    "FACTURA": current_factura,
                    "GLOBAL_TOTAL_BPRO": factura_total_global, # Guardamos el total global para referencia
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
# 4. LÓGICA MATEMÁTICA "SUBSET SUM" (Con Válvula de Seguridad)
# ======================================================
def subset_exact_safe(items, objetivo, max_items=15):
    """
    Intenta encontrar qué combinación de 'items' suma exactamente el 'objetivo'.
    SEGURIDAD: Si hay más de 'max_items', aborta para evitar colgar el servidor.
    """
    # 1. Válvula de Seguridad: Si es una factura gigante, no arriesgar
    if len(items) > max_items:
        return None 
        
    tolerancia = Decimal("0.05") # 5 centavos de tolerancia
    
    # 2. Fuerza Bruta Controlada
    for r in range(1, len(items) + 1):
        for combo in itertools.combinations(items, r):
            suma = sum(x["val"] for x in combo)
            if abs(suma - objetivo) <= tolerancia:
                return combo # ¡Match encontrado!
    return None

def apply_vendor_logic_safe(df, mapa_vendedores):
    """
    Orquesta la asignación.
    - Si es 1 vendedor -> Asigna directo.
    - Si son varios -> Usa matemáticas.
    - Si falla matemáticas o es muy grande -> Asigna al mayoritario (Plan B).
    """
    df["VENDEDOR"] = "Sin Vendedor"

    # Procesamos factura por factura (Grupo por Grupo)
    for factura, g in df.groupby("FACTURA"):
        indices = g.index
        candidatos = mapa_vendedores.get(factura, [])

        if not candidatos: continue

        # Determinar signo basado en BPro Col G
        total_global = g["GLOBAL_TOTAL_BPRO"].iloc[0]
        es_pos = total_global >= 0

        # Crear lista de metas (Solo tomamos a quienes tengan saldo en la columna correcta)
        metas = []
        for c in candidatos:
            monto = c["POS"] if es_pos else c["NEG"]
            if abs(monto) > 0: 
                metas.append({"VENDEDOR": c["VENDEDOR"], "META": monto})
        
        if not metas: continue

        # CASO 1: Un solo vendedor (95% de los casos) -> Asignación Directa Rápida
        if len(metas) == 1:
            df.loc[indices, "VENDEDOR"] = metas[0]["VENDEDOR"]
            continue

        # CASO 2: Factura Compartida -> Intentar Match Matemático
        items = [{"idx": idx, "val": g.at[idx, "VENTA SUBTOTAL"]} for idx in indices]
        usados = set()
        
        # Ordenamos metas: Intentar satisfacer primero al que pide más dinero (Mejor probabilidad de éxito)
        metas.sort(key=lambda x: abs(x["META"]), reverse=True)
        
        fallo_matematico = False
        
        for meta in metas:
            disponibles = [x for x in items if x["idx"] not in usados]
            
            # Llamamos a la función segura
            sol = subset_exact_safe(disponibles, meta["META"], max_items=15)

            if sol:
                for it in sol:
                    df.at[it["idx"], "VENDEDOR"] = meta["VENDEDOR"]
                    usados.add(it["idx"])
            else:
                fallo_matematico = True
                
        # PLAN DE RESPALDO (Si falló la matemática o activó la válvula de seguridad)
        if fallo_matematico:
            # Asignar todo lo sobrante al vendedor con la meta más grande (Lider)
            # Esto evita dejar filas "Sin Vendedor" y evita crashes.
            vendedor_lider = metas[0]["VENDEDOR"] 
            for idx in indices:
                if df.at[idx, "VENDEDOR"] == "Sin Vendedor":
                    df.at[idx, "VENDEDOR"] = vendedor_lider

    return df

# ======================================================
# 5. EJECUCIÓN PRINCIPAL (MAIN)
# ======================================================
# ======================================================
# INTERFAZ DE CARGA
# ======================================================
uid = st.session_state['id_limpieza']

# Selectores de archivo
col1, col2, col3 = st.columns(3)
with col1: file_bpro = st.file_uploader("1. Reporte BPro (.xls)", type=["xls"], key=f"b_{uid}")
with col2: file_vend = st.file_uploader("2. Reporte Vendedores (.xls/x)", type=["xls", "xlsx"], key=f"v_{uid}")
with col3: file_cat = st.file_uploader("3. Catálogo Partes (.xlsx)", type=["xls", "xlsx"], key=f"c_{uid}")

# ======================================================
# EJECUCIÓN PRINCIPAL CON BARRA DE PROGRESO
# ======================================================
if st.button("🚀 GENERAR BASE MAESTRA"):
    if not (file_bpro and file_vend and file_cat):
        st.warning("⚠️ Por favor, sube los 3 archivos para continuar.")
    else:
        # 1. INICIALIZAR BARRA DE PROGRESO
        progress_text = "Iniciando el motor de procesamiento..."
        my_bar = st.progress(0, text=progress_text)

        try:
            # --- ETAPA 1: LEER VENDEDORES (10%) ---
            my_bar.progress(10, text="📖 Leyendo y analizando archivo de Vendedores...")
            mapa = process_vendedores_file(file_vend)
            
            # --- ETAPA 2: LEER BPRO (30%) ---
            my_bar.progress(30, text="🏭 Procesando archivo BPro y detectando facturas...")
            df = process_main_bpro(file_bpro)
            
            if df.empty:
                st.error("Error: El archivo BPro parece estar vacío o con formato irreconocible.")
                my_bar.empty() # Borrar barra si falla
            else:
                df["SUCURSAL"] = sucursal

                # --- ETAPA 3: CÁLCULO MATEMÁTICO (Esta es la pesada - 60%) ---
                my_bar.progress(50, text="🧮 Realizando cruce matemático inteligente (Subset Sum)...")
                
                # Aquí ocurre la magia
                df = apply_vendor_logic_safe(df, mapa)
                
                # --- ETAPA 4: CATÁLOGO Y LIMPIEZA (80%) ---
                my_bar.progress(80, text="📚 Cruzando con Catálogo y aplicando reglas de negocio...")
                
                df_cat = pd.read_excel(file_cat)
                df["NP"] = df["NP"].astype(str).str.strip()
                df_cat["NP"] = df_cat["NP"].astype(str).str.strip()
                
                df_final = pd.merge(df, df_cat[['NP', 'LINEA', 'CATEGORIA']], 
                                  on='NP', how='left', suffixes=('_orig', '_new'))
                
                # Rellenos de Catálogo
                mask_nocat = df_final['CATEGORIA'].isna()
                df_final.loc[mask_nocat, 'DESCR'] = 'X DEFINIR'
                df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('X DEFINIR')
                df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                
                # --- ETAPA 5: FORMATO FINAL (95%) ---
                my_bar.progress(95, text="✨ Aplicando formatos finales y generando Excel...")
                
                # Correcciones Numéricas
                for col in ["CANTIDAD", "COSTO DE VENTA", "VENTA SUBTOTAL", "UTILIDAD"]:
                    df_final[col] = pd.to_numeric(df_final[col], errors="coerce").fillna(0)

                # Corrección Signos (Devoluciones)
                mask_err = (df_final["CANTIDAD"] < 0) & (df_final["COSTO DE VENTA"] > 0)
                df_final.loc[mask_err, "COSTO DE VENTA"] *= -1
                df_final.loc[mask_err, "UTILIDAD"] = df_final.loc[mask_err, "VENTA SUBTOTAL"] - df_final.loc[mask_err, "COSTO DE VENTA"]

                # Fechas
                df_final["FECHA_DT"] = pd.to_datetime(df_final["FECHA"], dayfirst=True, errors="coerce")
                if anio_texto.isdigit():
                    df_final = df_final[df_final["FECHA_DT"].dt.year == int(anio_texto)]

                df_final["año"] = df_final["FECHA_DT"].dt.year
                df_final["mes"] = df_final["FECHA_DT"].apply(get_mes_espanol)
                df_final["FECHA"] = df_final["FECHA_DT"].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "Sin Fecha")
                
                # Ordenar
                df_final = df_final.sort_values(by=['FECHA_DT', 'CLIENTE'])

                # Selección Final
                cols_fin = ["SUCURSAL", "CLIENTE", "FACTURA", "VENDEDOR", "NP", "DESCR", "CATEGORIA", "LINEA", "FECHA", "año", "mes", "TIPO", "ALMACEN", "CANTIDAD", "COSTO", "PRECIO", "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"]
                for c in cols_fin:
                    if c not in df_final.columns: df_final[c] = 0
                df_final = df_final[cols_fin]

                # --- COMPLETADO (100%) ---
                my_bar.progress(100, text="✅ ¡Proceso Completado!")
                time.sleep(1) # Pequeña pausa visual
                my_bar.empty() # Quitamos la barra para que se vea limpio

                st.success(f"¡Base generada exitosamente! ({len(df_final)} registros procesados)")
                st.dataframe(df_final.head(10))
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Datos')
                
                st.download_button(f"📥 Descargar {sucursal}", data=buffer.getvalue(), 
                                 file_name=f"{sucursal}_Ventas_{anio_texto}_MASTER.xlsx",
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"Error Técnico Inesperado: {e}")
            # Si tienes la barra activa y falla, es bueno limpiarla o mostrar error
            if 'my_bar' in locals():
                my_bar.empty()

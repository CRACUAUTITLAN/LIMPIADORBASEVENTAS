import streamlit as st
import pandas as pd
import io
import datetime
import re
import requests
from google.oauth2 import service_account
import google.auth.transport.requests

# ==========================================
# ⚙️ FUNCIONES AUXILIARES
# ==========================================
def normalizar_hora(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
        return "00:00:00"
    if isinstance(val, (pd.Timestamp, datetime.time, datetime.datetime)):
        try:
            return val.strftime('%H:%M:%S')
        except:
            pass
    val_str = str(val).strip().lower()
    if " " in val_str and (":" in val_str.split(" ")[1] or "m" in val_str):
        partes_espacio = val_str.split(" ")
        for parte in partes_espacio:
            if ":" in parte:
                val_str = parte
                if "a" in val_str or "p" in val_str:
                    val_str += " pm" if "p" in str(val).lower() else " am"
                break
    es_pm = False
    if "p" in val_str:
        es_pm = True
    hora_limpia = re.sub(r'[^0-9:]', '', val_str)
    if not hora_limpia or ":" not in hora_limpia:
        try:
            total_seconds = int(float(val_str) * 86400)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        except:
            return "00:00:00"
    partes = hora_limpia.split(":")
    try:
        h = int(partes[0])
        m = int(partes[1]) if len(partes) > 1 else 0
        s = int(partes[2]) if len(partes) > 2 else 0
        if es_pm and h < 12:
            h += 12
        elif not es_pm and h == 12:
            h = 0
        if h >= 24 or m >= 60 or s >= 60:
            return "00:00:00"
        return f"{h:02d}:{m:02d}:{s:02d}"
    except:
        return "00:00:00"

# ==========================================
# 🤖 AUTENTICACIÓN DEL ROBOT
# ==========================================
@st.cache_resource
def get_auth_headers():
    try:
        gcp_creds = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            gcp_creds,
            scopes=['https://www.googleapis.com/auth/drive.readonly']
        )
        auth_req = google.auth.transport.requests.Request()
        creds.refresh(auth_req) 
        return {"Authorization": f"Bearer {creds.token}"}
    except Exception as e:
        st.error(f"⚠️ Error al conectar el Robot con Google: {e}")
        st.stop()

def show():
    st.title("📊 Master App: Fill Rate & Análisis What-If (Kardex Virtual)")
    st.markdown("Extrae los datos en la nube, estandariza y simula el consumo del stock fila por fila.")

    # ==========================================
    # SECCIÓN 1 Y 2: GENERACIÓN BASE MADRE
    # ==========================================
    st.markdown("### 📚 1. Catálogo Maestro de Refacciones")
    archivo_catalogo = st.file_uploader("Cargar Catálogo Maestro", type=["xlsx", "xls", "csv"], key="cat_maestro")

    st.markdown("---")
    st.markdown("### 🗂️ 2. Extracción Segura en la Nube (Base Madre)")

    CONSIGNAS_LINKS = {
        "ALMACEN AFN": "https://docs.google.com/spreadsheets/d/1eDMQxTNrAJ3Ak-RLASoQyYmQP_UpubsLSDT3m-K_sIg/export?format=csv&gid=0",
        "BISONTE SLP": "https://docs.google.com/spreadsheets/d/1KYCIqdU95WtI_wSl48NYnZHunq93_blwTBr4lJY1MQI/export?format=csv&gid=0",
        "BISONTE TEPOTZOTLAN": "https://docs.google.com/spreadsheets/d/1ZqpyM-2jH2flO389_OFr25ac_3RAZvtwxHuOAJSrQMI/export?format=csv&gid=0",
        "CULVERT": "https://docs.google.com/spreadsheets/d/1G0KnKiE_EklbNFSUJpny3yOrJoLrqGiodDbPWJg1JD0/export?format=csv&gid=0",
        "IDEALEASE": "https://docs.google.com/spreadsheets/d/1ikW-McoDPbAcHZQTf6VxwyVjg7IzO1JUX5DIBRUGRk0/export?format=csv&gid=0",
        "ALM. LIPU": "https://docs.google.com/spreadsheets/d/1NoGkLlXbNVDesKL8l5KB11BuVBtW4wsMMCIi80FqlQk/export?format=csv&gid=0",
        "ALM. MYM": "https://docs.google.com/spreadsheets/d/1GhvVOEJIjSXT9m-pEnhpM2sL5aI480Axg4ImpcUgH_U/export?format=csv&gid=0",
        "TDR": "https://docs.google.com/spreadsheets/d/18PMsSuZgKmxhDnfzUndKgFhzXf9iR2ZDKKGsdBp2rMQ/export?format=csv&gid=0",
        "TUMSA": "https://docs.google.com/spreadsheets/d/1Hb4gQ5mLdkxyWKBAz--cyPiRxyMbDdHYm-ROd5NxLZI/export?format=csv&gid=0",
        "ALM. UTEP": "https://docs.google.com/spreadsheets/d/1HuV2-EGIb0kfpuDdAVLCh1mzauBifc28nMzjG7h_Cx4/export?format=csv&gid=0",
        "ALM. UTEP SAN LUIS": "https://docs.google.com/spreadsheets/d/1tdRJEgnmb3Mg8kOhXEAvgPsMQLz60UE6nhLWLV1xk9o/export?format=csv&gid=0"
    }

    columnas_finales = [
        "AGENCIA", "ALMACEN", "NP", "DESCR", 
        "FECHA SOLICITUD (DIA/MES/AÑO)", "HORA SOLICITUD", 
        "FECHA ENTREGA (DIA/MES/AÑO)", "HORA ENTREGA (Formato 24 hrs)", 
        "DIAS ENTREGA", "CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "ESTATUS"
    ]

    def get_agencia(nombre_almacen):
        if nombre_almacen in ["ALM. LIPU", "ALM. MYM", "ALM. UTEP", "ALM. UTEP SAN LUIS", "IDEALEASE"]:
            return "CUAUTITLAN"
        return "TULTITLAN"

    if st.button("🚀 Extraer Datos con el Robot y Generar Base", type="primary"):
        if not archivo_catalogo:
            st.warning("⚠️ Por favor sube primero el Catálogo Maestro de Refacciones.")
            return

        with st.spinner("🤖 El Robot se está conectando a Google Drive..."):
            headers = get_auth_headers()
            try:
                if archivo_catalogo.name.endswith('.csv'):
                    df_cat = pd.read_csv(archivo_catalogo)
                else:
                    df_cat = pd.read_excel(archivo_catalogo)
                df_cat.columns = [str(col).strip().upper() for col in df_cat.columns]
                col_descr_cat = "DESCRIPCION" if "DESCRIPCION" in df_cat.columns else "DESCRIPCIÓN"
                df_cat = df_cat[['NP', col_descr_cat]].drop_duplicates(subset=['NP'])
                df_cat['NP'] = df_cat['NP'].astype(str).str.strip()
                df_cat = df_cat.rename(columns={col_descr_cat: "DESCR_OFICIAL"})
            except Exception as e:
                st.error(f"❌ Error al procesar el Catálogo Maestro: {e}")
                return

            dataframes_limpios = []
            errores = []
            barra = st.progress(0)
            total_links = len(CONSIGNAS_LINKS)

            for i, (almacen, url) in enumerate(CONSIGNAS_LINKS.items()):
                agencia = get_agencia(almacen)
                try:
                    response = requests.get(url, headers=headers)
                    if response.status_code == 200:
                        df = pd.read_csv(io.StringIO(response.text), skiprows=11)
                        df.columns = [str(c).strip() for c in df.columns]
                        for col in df.columns:
                            col_upper = col.upper()
                            if "PARTE" in col_upper or "NP" == col_upper:
                                df = df.rename(columns={col: "NP"})
                            elif "DESCRIP" in col_upper:
                                df = df.rename(columns={col: "DESCR"})
                            elif "FECHA SOL" in col_upper:
                                df = df.rename(columns={col: "FECHA SOLICITUD (DIA/MES/AÑO)"})
                            elif "FECHA ENT" in col_upper:
                                df = df.rename(columns={col: "FECHA ENTREGA (DIA/MES/AÑO)"})
                            elif "HORA ENT" in col_upper:
                                df = df.rename(columns={col: "HORA ENTREGA (Formato 24 hrs)"})
                        if "NP" in df.columns:
                            df = df.dropna(subset=["NP"])
                        if "CONSIGNA" in df.columns:
                            df["CONSIGNA_UPPER"] = df["CONSIGNA"].astype(str).str.upper()
                            if almacen == "ALM. LIPU":
                                df = df[df["CONSIGNA_UPPER"].str.contains("LIPU", na=False)]
                            elif almacen == "ALM. UTEP":
                                df = df[df["CONSIGNA_UPPER"].str.contains("UTEP", na=False)]
                        df["AGENCIA"] = agencia
                        df["ALMACEN"] = almacen
                        columnas_existentes = [col for col in columnas_finales if col in df.columns]
                        for col_faltante in set(columnas_finales) - set(columnas_existentes):
                            df[col_faltante] = None
                        df_limpio = df[columnas_finales].copy()
                        dataframes_limpios.append(df_limpio)
                    else:
                        errores.append(f"❌ Acceso Denegado a '{almacen}'. HTTP {response.status_code}")
                except Exception as e:
                    errores.append(f"❌ Error al extraer '{almacen}': {str(e)}")
                barra.progress((i + 1) / total_links)

            for err in errores:
                st.error(err)

            if dataframes_limpios:
                df_master = pd.concat(dataframes_limpios, ignore_index=True)
                df_master = df_master.dropna(subset=["NP"])
                df_master["NP"] = df_master["NP"].astype(str).str.strip()
                df_master = df_master[df_master["NP"] != ""]
                df_master = df_master[df_master["NP"].str.lower() != "nan"]
                df_master = df_master.merge(df_cat, on='NP', how='left')
                df_master['DESCR'] = df_master['DESCR_OFICIAL'].fillna("X DEFINIR")
                df_master = df_master.drop(columns=['DESCR_OFICIAL'])

                columnas_num = ["CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "DIAS ENTREGA"]
                for col in columnas_num:
                    df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)
                    df_master[col] = df_master[col].astype(int)

                df_master["HORA SOLICITUD"] = df_master["HORA SOLICITUD"].apply(normalizar_hora)
                df_master["HORA ENTREGA (Formato 24 hrs)"] = df_master["HORA ENTREGA (Formato 24 hrs)"].apply(normalizar_hora)

                for col_fecha in ["FECHA SOLICITUD (DIA/MES/AÑO)", "FECHA ENTREGA (DIA/MES/AÑO)"]:
                    df_master[col_fecha] = pd.to_datetime(df_master[col_fecha], dayfirst=True, errors='coerce')
                    if col_fecha == "FECHA ENTREGA (DIA/MES/AÑO)":
                        df_master[col_fecha] = df_master[col_fecha].fillna(df_master["FECHA SOLICITUD (DIA/MES/AÑO)"])
                    df_master[col_fecha] = df_master[col_fecha].fillna(pd.Timestamp('2026-01-01'))
                    df_master[col_fecha] = df_master[col_fecha].dt.strftime('%Y-%m-%d')

                st.success(f"✅ ¡Base Madre extraída con éxito! ({len(df_master)} registros)")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_master.to_excel(writer, index=False, sheet_name='Fact_Fill_Rate')
                st.download_button(
                    label="📥 Descargar Base Madre Normalizada (Para usar abajo 👇)",
                    data=buffer.getvalue(),
                    file_name="BASE_MADRE_FILL_RATE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

    st.markdown("---")
    # ==========================================
    # SECCIÓN 3: SIMULADOR WHAT-IF (KARDEX DINÁMICO REGLA DE TIEMPO)
    # ==========================================
    st.markdown("### 🔮 3. Análisis What-If (Simulador de Traspasos Inteligente)")
    st.info("Solo se evaluará el rescate en las solicitudes donde la Fecha de Entrega superó la de Solicitud (Entregas Tardías).")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        archivo_fill_rate = st.file_uploader("1️⃣ Base Madre Fill Rate (.xlsx)", type=["xlsx"], key="fill_rate_file")
    with col2:
        archivo_traspasos = st.file_uploader("2️⃣ Análisis Consignas (Reporte Sábado)", type=["xlsx"], key="traspasos_file")
    with col3:
        fecha_reporte = st.date_input("📅 Fecha del Reporte (Sábado):")

    if st.button("🎯 Simular Kardex y Fill Rate", type="primary", use_container_width=True):
        if not archivo_fill_rate or not archivo_traspasos:
            st.warning("⚠️ Necesitas cargar ambos archivos para realizar la simulación.")
            return
            
        with st.spinner("🧠 Evaluando entregas tardías y reconstruyendo Kardex..."):
            try:
                # 1. Cargar Fill Rate y crear línea de tiempo
                df_fill = pd.read_excel(archivo_fill_rate)
                df_fill['FECHA_SOL_DT'] = pd.to_datetime(df_fill['FECHA SOLICITUD (DIA/MES/AÑO)'], errors='coerce')
                df_fill['FECHA_ENT_DT'] = pd.to_datetime(df_fill['FECHA ENTREGA (DIA/MES/AÑO)'], errors='coerce')
                
                df_fill['FECHA_HORA_ORDEN'] = pd.to_datetime(df_fill['FECHA SOLICITUD (DIA/MES/AÑO)'].astype(str) + ' ' + df_fill['HORA SOLICITUD'].astype(str), errors='coerce')
                df_fill = df_fill.dropna(subset=['FECHA_HORA_ORDEN']) 
                
                df_fill['CANTIDAD SOLICITADA'] = pd.to_numeric(df_fill['CANTIDAD SOLICITADA'], errors='coerce').fillna(0).astype(float)
                df_fill['CANTIDAD ENTREGADA'] = pd.to_numeric(df_fill['CANTIDAD ENTREGADA'], errors='coerce').fillna(0).astype(float)
                
                # Filtro de fecha
                fecha_corte = pd.to_datetime(fecha_reporte).date()
                df_fill = df_fill[df_fill['FECHA_HORA_ORDEN'].dt.date > fecha_corte].copy()
                
                if df_fill.empty:
                    st.error("❌ No hay solicitudes posteriores a la fecha seleccionada.")
                    return

                df_fill = df_fill.sort_values('FECHA_HORA_ORDEN')

                # =========================================================
                # 2. CARGA INTELIGENTE DE TRASPASOS
                # =========================================================
                inventario_virtual = {}
                xls_traspasos = pd.ExcelFile(archivo_traspasos)
                
                almacenes_en_fill_rate = df_fill['ALMACEN'].dropna().unique()
                
                for almacen_fr in almacenes_en_fill_rate:
                    alm_limpio = str(almacen_fr).strip().upper()
                    
                    hoja_encontrada = None
                    for sheet in xls_traspasos.sheet_names:
                        if sheet.strip().upper() == alm_limpio:
                            hoja_encontrada = sheet
                            break
                            
                    if hoja_encontrada:
                        df_raw = pd.read_excel(xls_traspasos, sheet_name=hoja_encontrada, header=None)
                        header_idx = 0
                        for i, r in df_raw.iterrows():
                            if "N° DE PARTE" in str(r.values).upper() or "NP" in str(r.values).upper():
                                header_idx = i
                                break
                                
                        df_alm = pd.read_excel(xls_traspasos, sheet_name=hoja_encontrada, header=header_idx)
                        df_alm.columns = [str(c).strip().upper() for c in df_alm.columns]
                        
                        col_np = 'N° DE PARTE' if 'N° DE PARTE' in df_alm.columns else ('NP' if 'NP' in df_alm.columns else None)
                        if not col_np: continue
                        
                        for _, r in df_alm.iterrows():
                            np_val = str(r[col_np]).strip().upper()
                            if np_val in ('NAN', '', 'NONE'): continue
                            
                            demanda = str(r.get('DEMANDA', '')).strip().upper()
                            
                            try:
                                if demanda == 'ALTA': meta = float(r.get('ALTA (1.5)', 0))
                                elif demanda == 'MEDIA': meta = float(r.get('MEDIA (1)', 0))
                                else: meta = float(r.get('BAJA (0)', 0))
                            except:
                                meta = 0
                                
                            try:
                                existencia = float(r.get('INVENTARIO EXISTENCIA', 0))
                            except:
                                existencia = 0
                                
                            stock_arranque = max(meta, existencia)
                            inventario_virtual[(alm_limpio, np_val)] = stock_arranque

                # =========================================================
                # 3. KARDEX CON REGLA ESTRICTA DE TIEMPO
                # =========================================================
                lista_stock_sugerido = []
                lista_entrega_supuesta = []
                lista_fr_real = []
                lista_fr_supuesto = []
                
                for idx, row in df_fill.iterrows():
                    alm_key = str(row['ALMACEN']).strip().upper()
                    np_key = str(row['NP']).strip().upper()
                    cant_solicitada = row['CANTIDAD SOLICITADA']
                    cant_entregada = row['CANTIDAD ENTREGADA']
                    
                    llave = (alm_key, np_key)
                    stock_al_momento = inventario_virtual.get(llave, 0.0)

                    # Comprobar la regla de oro: ¿Fecha Entrega > Fecha Solicitud?
                    entregado_tarde = False
                    if pd.notna(row['FECHA_SOL_DT']) and pd.notna(row['FECHA_ENT_DT']):
                        if row['FECHA_ENT_DT'].date() > row['FECHA_SOL_DT'].date():
                            entregado_tarde = True
                    
                    # Guardar siempre la foto del stock sugerido al momento del evento
                    lista_stock_sugerido.append(stock_al_momento)

                    if entregado_tarde:
                        # ❌ FALLAMOS EN EL TIEMPO: Entra el Análisis What-If
                        fr_real = 0.0 
                        entrega_supuesta = min(cant_solicitada, stock_al_momento)
                        fr_supuesto = (entrega_supuesta / cant_solicitada) if cant_solicitada > 0 else 0.0
                        
                        # Drenamos del Kardex Virtual la cantidad que el modelo usó para rescatar
                        if llave in inventario_virtual:
                            inventario_virtual[llave] = max(0.0, stock_al_momento - entrega_supuesta)
                            
                    else:
                        # ✅ ENTREGADO EL MISMO DÍA: No intervenimos, copiamos tal cual
                        fr_real = min((cant_entregada / cant_solicitada), 1.0) if cant_solicitada > 0 else 0.0
                        entrega_supuesta = cant_entregada
                        fr_supuesto = fr_real
                        
                        # Drenamos del Kardex Virtual lo que físicamente SÍ entregaron
                        # para que el Kardex vaya cuadrado a la realidad
                        if llave in inventario_virtual:
                            inventario_virtual[llave] = max(0.0, stock_al_momento - cant_entregada)

                    # Guardar resultados
                    lista_entrega_supuesta.append(entrega_supuesta)
                    lista_fr_real.append(fr_real)
                    lista_fr_supuesto.append(fr_supuesto)

                # 4. Inyectar Nuevas Columnas a la Base
                df_fill['STOCK SUGERIDO'] = lista_stock_sugerido
                df_fill['ENTREGA SUPUESTA'] = lista_entrega_supuesta
                df_fill['FILL RATE'] = lista_fr_real
                df_fill['FILL RATE SUPUESTO'] = lista_fr_supuesto
                
                df_fill = df_fill.drop(columns=['FECHA_HORA_ORDEN', 'FECHA_SOL_DT', 'FECHA_ENT_DT'])

                promedio_fr_real = df_fill['FILL RATE'].mean() * 100
                promedio_fr_sup = df_fill['FILL RATE SUPUESTO'].mean() * 100
                
                st.success(f"✅ ¡Kardex simulado correctamente! Evaluación aplicada únicamente a las entregas tardías.")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("Fill Rate Real (Promedio)", f"{promedio_fr_real:.1f}%")
                m2.metric("Fill Rate Supuesto", f"{promedio_fr_sup:.1f}%", f"+{(promedio_fr_sup - promedio_fr_real):.1f}%")
                m3.metric("Líneas Analizadas", len(df_fill))

                buffer_bi = io.BytesIO()
                with pd.ExcelWriter(buffer_bi, engine='openpyxl') as writer:
                    df_fill.to_excel(writer, index=False, sheet_name='Kardex_WhatIf')
                
                st.download_button(
                    label="📊 Descargar Base Kardex para Power BI",
                    data=buffer_bi.getvalue(),
                    file_name=f"BASE_FILL_RATE_WHAT_IF_{fecha_corte}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.markdown("#### 🏆 Vista de Consumo del Kardex (Fila por Fila)")
                cols_mostrar = ['ALMACEN', 'NP', 'FECHA SOLICITUD (DIA/MES/AÑO)', 'FECHA ENTREGA (DIA/MES/AÑO)', 'CANTIDAD SOLICITADA', 'CANTIDAD ENTREGADA', 'STOCK SUGERIDO', 'ENTREGA SUPUESTA', 'FILL RATE', 'FILL RATE SUPUESTO']
                st.dataframe(df_fill[cols_mostrar].head(30))
                
            except Exception as e:
                st.error(f"❌ Ocurrió un error en la simulación: {e}")

if __name__ == "__main__":
    show()

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
    """
    Toma cualquier formato de hora de Excel (Texto, AM/PM, decimal o datetime)
    y lo convierte a un formato ESTRICTO de 24 horas: 'HH:MM:SS'.
    """
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
    """Obtiene el token de seguridad del Robot para entrar a Google Drive"""
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
    st.title("📊 Master App: Fill Rate & Análisis What-If")
    st.markdown("Extrae los datos en la nube, estandariza columnas y simula el Fill Rate Supuesto cruzando con el plan de traspasos.")

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
            st.warning("⚠️ Por favor sube primero el Catálogo Maestro de Refacciones en el paso 1.")
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

                st.success(f"✅ ¡Base Madre extraída y blindada con éxito! ({len(df_master)} registros)")
                
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
    # SECCIÓN 3: SIMULADOR WHAT-IF (BASADO EN PIEZAS Y FILL RATE REAL)
    # ==========================================
    st.markdown("### 🔮 3. Análisis What-If (Simulador de Traspasos por Sucursal)")
    st.info("Calcula el Fill Rate por PIEZAS. Las piezas sugeridas por el modelo salvarán entregas parciales o tardías.")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        archivo_fill_rate = st.file_uploader("1️⃣ Base Madre Fill Rate (.xlsx)", type=["xlsx"], key="fill_rate_file")
    with col2:
        archivo_traspasos = st.file_uploader("2️⃣ Análisis Consignas (Reporte Sábado)", type=["xlsx"], key="traspasos_file")
    with col3:
        fecha_reporte = st.date_input("📅 Fecha del Reporte (Sábado):", help="Solo se evaluarán solicitudes posteriores a esta fecha.")

    if st.button("🎯 Simular Fill Rate Supuesto", type="primary", use_container_width=True):
        if not archivo_fill_rate or not archivo_traspasos:
            st.warning("⚠️ Necesitas cargar ambos archivos para realizar la simulación.")
            return
            
        with st.spinner("🧠 Calculando volúmenes de piezas, rescates e inyectando traspasos..."):
            try:
                # 1. Cargar Fill Rate y crear línea de tiempo
                df_fill = pd.read_excel(archivo_fill_rate)
                df_fill['FECHA_HORA_ORDEN'] = pd.to_datetime(df_fill['FECHA SOLICITUD (DIA/MES/AÑO)'].astype(str) + ' ' + df_fill['HORA SOLICITUD'].astype(str), errors='coerce')
                
                # Blindaje de Fechas Nulas
                df_fill = df_fill.dropna(subset=['FECHA_HORA_ORDEN']) 

                # Blindaje Columna de Cantidades
                df_fill['CANTIDAD SOLICITADA'] = pd.to_numeric(df_fill['CANTIDAD SOLICITADA'], errors='coerce').fillna(0)
                
                # 2. FILTRO MAESTRO DE FECHA
                fecha_corte = pd.to_datetime(fecha_reporte).date()
                df_fill = df_fill[df_fill['FECHA_HORA_ORDEN'].dt.date > fecha_corte].copy()
                
                if df_fill.empty:
                    st.error("❌ No hay solicitudes de Fill Rate posteriores a la fecha seleccionada. Revisa la fecha.")
                    return

                df_fill = df_fill.sort_values('FECHA_HORA_ORDEN')
                
                # 3. Matemática de Piezas Reales (El Fill Rate Hoy)
                def calcular_entregado_dia_0(row):
                    try:
                        dias = float(row.get('DIAS ENTREGA', 0))
                        if dias == 0:
                            return float(row.get('CANTIDAD ENTREGADA', 0))
                        else:
                            return 0.0
                    except:
                        return 0.0
                        
                df_fill['CANT_REAL_DIA_0'] = df_fill.apply(calcular_entregado_dia_0, axis=1)
                df_fill['PIEZAS_FALTANTES_REALES'] = df_fill['CANTIDAD SOLICITADA'].astype(float) - df_fill['CANT_REAL_DIA_0']

                # 4. Cargar Base de Traspasos (El Fondo de Rescate)
                stock_rescate_modelo = {}
                xls_traspasos = pd.ExcelFile(archivo_traspasos)
                
                almacenes_objetivo = [
                    "ALM. LIPU", "ALM. MYM", "ALM. UTEP", "ALM. UTEP SAN LUIS", 
                    "ALMACEN AFN", "BISONTE SLP", "BISONTE TEPOTZOTLAN", 
                    "CULVERT", "IDEALEASE", "TDR", "TUMSA"
                ]

                # Leer hoja por hoja
                for almacen in almacenes_objetivo:
                    if almacen in xls_traspasos.sheet_names:
                        df_alm = pd.read_excel(xls_traspasos, sheet_name=almacen)
                        
                        header_idx = 0
                        for i, row in df_alm.iterrows():
                            if "N° DE PARTE" in str(row.values).upper() or "NP" in str(row.values).upper():
                                header_idx = i
                                break
                        
                        if header_idx > 0:
                            df_alm = pd.read_excel(xls_traspasos, sheet_name=almacen, header=header_idx)
                        
                        for _, row in df_alm.iterrows():
                            np_val = str(row.get('N° DE PARTE', row.get('NP', ''))).strip()
                            if np_val in ('nan', '', 'None'): continue
                            
                            demanda = str(row.get('DEMANDA', '')).strip().upper()
                            
                            try:
                                if demanda == 'ALTA': meta = float(row.get('ALTA (1.5)', 0))
                                elif demanda == 'MEDIA': meta = float(row.get('MEDIA (1)', 0))
                                else: meta = float(row.get('BAJA (0)', 0))
                            except:
                                meta = 0
                                
                            try:
                                existencia = float(row.get('INVENTARIO EXISTENCIA', 0))
                            except:
                                existencia = 0
                                
                            piezas_traspasadas = max(0.0, meta - existencia)
                            stock_rescate_modelo[(almacen, np_val)] = piezas_traspasadas
                            
                # 5. La Simulación (Rescatando faltantes)
                piezas_salvadas_list = []
                cant_supuesta_dia_0 = []
                
                for idx, row in df_fill.iterrows():
                    np_val = str(row['NP']).strip()
                    almacen = str(row['ALMACEN']).strip()
                    faltante = row['PIEZAS_FALTANTES_REALES']
                    real_dia_0 = row['CANT_REAL_DIA_0']
                    
                    llave = (almacen, np_val)
                    disp_rescate = stock_rescate_modelo.get(llave, 0)
                    
                    # AQUÍ ESTÁ LA CORRECCIÓN EXACTA QUE INCLUYE LOS FALTANTES "NO SALVADOS"
                    if faltante > 0 and disp_rescate > 0:
                        piezas_salvadas = min(faltante, disp_rescate)
                        stock_rescate_modelo[llave] -= piezas_salvadas
                    else:
                        piezas_salvadas = 0.0 # Se registra la fila, pero con 0 piezas salvadas.

                    piezas_salvadas_list.append(piezas_salvadas)
                    cant_supuesta_dia_0.append(real_dia_0 + piezas_salvadas)

                # Insertar columnas para Power BI
                df_fill['PIEZAS SALVADAS POR MODELO'] = piezas_salvadas_list
                df_fill['CANT_SUPUESTA_DIA_0'] = cant_supuesta_dia_0
                
                # Estatus para tabla resumen
                df_fill['ESTATUS SUPUESTO'] = df_fill.apply(
                    lambda x: "✅ SALVADO AL 100%" if x['PIEZAS SALVADAS POR MODELO'] == x['PIEZAS_FALTANTES_REALES'] and x['PIEZAS_FALTANTES_REALES'] > 0
                    else ("⚠️ SALVADO PARCIAL" if x['PIEZAS SALVADAS POR MODELO'] > 0 else ("✅ BIEN (Mismo día)" if x['PIEZAS_FALTANTES_REALES'] == 0 else "❌ NO SALVADO")), axis=1
                )

                df_fill = df_fill.drop(columns=['FECHA_HORA_ORDEN'])

                # ==========================================
                # MÉTRICAS ESTRICTAS POR PIEZAS
                # ==========================================
                total_piezas_solicitadas = df_fill['CANTIDAD SOLICITADA'].sum()
                total_piezas_entregadas_reales = df_fill['CANT_REAL_DIA_0'].sum()
                total_piezas_salvadas = sum(piezas_salvadas_list)
                total_piezas_entregadas_supuesto = sum(cant_supuesta_dia_0)
                
                fr_real = (total_piezas_entregadas_reales / total_piezas_solicitadas * 100) if total_piezas_solicitadas > 0 else 0
                fr_supuesto = (total_piezas_entregadas_supuesto / total_piezas_solicitadas * 100) if total_piezas_solicitadas > 0 else 0
                
                st.success(f"✅ ¡Simulación Completada! (Basada en el volumen de {int(total_piezas_solicitadas)} piezas solicitadas)")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("Fill Rate Real (Piezas)", f"{fr_real:.1f}%", f"{int(total_piezas_entregadas_reales)} pzs a tiempo")
                m2.metric("Fill Rate Supuesto", f"{fr_supuesto:.1f}%", f"+{(fr_supuesto - fr_real):.1f}% de Mejora")
                m3.metric("Piezas Salvadas por tu Modelo", f"{int(total_piezas_salvadas)} piezas")

                # Botón de Descarga para Power BI
                buffer_bi = io.BytesIO()
                with pd.ExcelWriter(buffer_bi, engine='openpyxl') as writer:
                    df_fill.to_excel(writer, index=False, sheet_name='Simulacion_WhatIf_Piezas')
                
                st.download_button(
                    label="📊 Descargar Base Lista para Power BI (Simulada por Piezas)",
                    data=buffer_bi.getvalue(),
                    file_name=f"BASE_FILL_RATE_WHAT_IF_{fecha_corte}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                # Tabla resumen para comprobar la magia
                st.markdown("#### 🏆 Resumen de Piezas Rescatadas")
                df_rescatados = df_fill[df_fill['PIEZAS SALVADAS POR MODELO'] > 0]
                st.dataframe(df_rescatados[['ALMACEN', 'NP', 'CANTIDAD SOLICITADA', 'CANT_REAL_DIA_0', 'PIEZAS_FALTANTES_REALES', 'PIEZAS SALVADAS POR MODELO', 'ESTATUS SUPUESTO']])
                
            except Exception as e:
                st.error(f"❌ Ocurrió un error en la simulación: {e}")

if __name__ == "__main__":
    show()

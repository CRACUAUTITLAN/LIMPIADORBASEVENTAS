import streamlit as st
import pandas as pd
import io
import datetime
import re
import requests
from google.oauth2 import service_account
import google.auth.transport.requests

def normalizar_hora(val):
    """
    Toma cualquier formato de hora de Excel (Texto, AM/PM, decimal o datetime)
    y lo convierte a un formato ESTRICTO de 24 horas: 'HH:MM:SS'.
    Si está vacío o tiene un error crónico, devuelve '00:00:00'.
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
    st.title("📊 Generador Base Madre Automático: Fill Rate")
    st.markdown("Extrae los datos de Google Sheets usando el Robot seguro, estandariza columnas y genera la base final para Power BI.")

    st.markdown("### 📚 1. Catálogo Maestro de Refacciones")
    archivo_catalogo = st.file_uploader("Cargar Catálogo Maestro", type=["xlsx", "xls", "csv"], key="cat_maestro")

    st.markdown("---")
    st.markdown("### 🗂️ 2. Extracción Segura en la Nube")

    # Links automatizados
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

    # Orden final estricto e inalterable para Power BI
    columnas_finales = [
        "AGENCIA", "ALMACEN", "NP", "DESCR", "FECHA SOLICITUD (DIA/MES/AÑO)", 
        "HORA SOLICITUD", "FECHA ENTREGA (DIA/MES/AÑO)", "HORA ENTREGA (Formato 24 hrs)", 
        "DIAS ENTREGA", "CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "ESTATUS",
        "VALE", "CONSIGNA", "OC TRAXION", "ENTRADA ALMACEN", "FACTURA", "COMENTARIOS", "ÉXITO DE ENTREGA"
    ]

    def get_agencia(nombre_almacen):
        if nombre_almacen in ["ALM. LIPU", "ALM. MYM", "ALM. UTEP", "ALM. UTEP SAN LUIS", "IDEALEASE"]:
            return "CUAUTITLAN"
        return "TULTITLAN"

    # ==========================================
    # PROCESAMIENTO DE DATOS
    # ==========================================
    if st.button("🚀 Extraer Datos con el Robot y Generar Base", type="primary"):
        if not archivo_catalogo:
            st.warning("⚠️ Por favor sube primero el Catálogo Maestro de Refacciones en el paso 1.")
            return

        with st.spinner("🤖 El Robot se está conectando a Google Drive..."):
            
            headers = get_auth_headers()

            # Leer Catálogo Maestro
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
                        
                        # LIMPIEZA DE NOMBRES DE COLUMNAS (Evita que NP salga en blanco)
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
                        
                        # =========================================================
                        # 🛡️ FILTROS ESTRICTOS EXCLUSIVOS PARA LIPU Y UTEP
                        # =========================================================
                        if "CONSIGNA" in df.columns:
                            df["CONSIGNA_UPPER"] = df["CONSIGNA"].astype(str).str.upper()
                            
                            if almacen == "ALM. LIPU":
                                df = df[df["CONSIGNA_UPPER"].str.contains("LIPU", na=False)]
                            elif almacen == "ALM. UTEP":
                                df = df[df["CONSIGNA_UPPER"].str.contains("UTEP", na=False)]
                                
                            df = df.drop(columns=["CONSIGNA_UPPER"])
                        # =========================================================
                        
                        df["AGENCIA"] = agencia
                        df["ALMACEN"] = almacen

                        columnas_existentes = [col for col in columnas_finales if col in df.columns]
                        for col_faltante in set(columnas_finales) - set(columnas_existentes):
                            df[col_faltante] = None

                        df_limpio = df[columnas_finales].copy()
                        dataframes_limpios.append(df_limpio)
                    else:
                        errores.append(f"❌ Acceso Denegado a '{almacen}': El Robot no tiene permisos. HTTP {response.status_code}")

                except Exception as e:
                    errores.append(f"❌ Error al extraer '{almacen}': {str(e)}")
                
                barra.progress((i + 1) / total_links)

            for err in errores:
                st.error(err)

            # 4. Consolidar Base Madre
            if dataframes_limpios:
                df_master = pd.concat(dataframes_limpios, ignore_index=True)
                
                # Quitar posibles nulos arrastrados en el NP
                df_master = df_master.dropna(subset=["NP"])
                df_master["NP"] = df_master["NP"].astype(str).str.strip()
                df_master = df_master[df_master["NP"] != ""]
                df_master = df_master[df_master["NP"].str.lower() != "nan"]

                df_master = df_master.merge(df_cat, on='NP', how='left')
                df_master['DESCR'] = df_master['DESCR_OFICIAL'].fillna("X DEFINIR")
                df_master = df_master.drop(columns=['DESCR_OFICIAL'])

                # BLINDAJES
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
                st.dataframe(df_master.head(15))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_master.to_excel(writer, index=False, sheet_name='Fact_Fill_Rate')
                
                st.download_button(
                    label="📥 Descargar Base Madre Normalizada",
                    data=buffer.getvalue(),
                    file_name="BASE_MADRE_FILL_RATE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

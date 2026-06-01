
import streamlit as st
import pandas as pd
import io
import datetime
import re
import requests
from google.oauth2 import service_account
import google.auth.transport.requests

def normalizar_hora(val):
    # ... [MANTÉN LA MISMA FUNCIÓN normalizar_hora QUE YA TENÍAMOS] ...
    pass # (Asegúrate de pegar el código completo de la función aquí)

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
        creds.refresh(auth_req) # Refresca y obtiene un token válido
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

    columnas_finales = [
        "AGENCIA", "ALMACEN", "NP", "DESCR", "CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", 
        "ESTATUS", "VALE", "CONSIGNA", "OC TRAXION", "ENTRADA ALMACEN", "FACTURA", 
        "COMENTARIOS", "FECHA SOLICITUD", "HORA SOLICITUD", "FECHA ENTREGA", 
        "HORA ENTREGA", "DIAS ENTREGA", "ÉXITO DE ENTREGA"
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
            
            # 1. Obtener la Llave Maestra del Robot
            headers = get_auth_headers()

            # 2. Leer Catálogo Maestro
            try:
                if archivo_catalogo.name.endswith('.csv'):
                    df_cat = pd.read_csv(archivo_catalogo)
                else:
                    df_cat = pd.read_excel(archivo_catalogo)
                
                df_cat.columns = [col.strip().upper() for col in df_cat.columns]
                col_descr_cat = "DESCRIPCION" if "DESCRIPCION" in df_cat.columns else "DESCRIPCIÓN"
                
                df_cat = df_cat[['NP', col_descr_cat]].drop_duplicates(subset=['NP'])
                df_cat['NP'] = df_cat['NP'].astype(str).str.strip()
                df_cat = df_cat.rename(columns={col_descr_cat: "DESCR_OFICIAL"})
                
            except Exception as e:
                st.error(f"❌ Error al procesar el Catálogo Maestro: {e}")
                return

            # 3. Extraer de Google Sheets vía API Privada
            dataframes_limpios = []
            errores = []
            barra = st.progress(0)
            total_links = len(CONSIGNAS_LINKS)

            for i, (almacen, url) in enumerate(CONSIGNAS_LINKS.items()):
                agencia = get_agencia(almacen)
                
                try:
                    # El Robot "toca la puerta" de Drive con su llave y descarga el CSV en memoria
                    response = requests.get(url, headers=headers)
                    
                    if response.status_code == 200:
                        # Leer el texto descargado ignorando la cabecera (11 filas)
                        df = pd.read_csv(io.StringIO(response.text), skiprows=11)
                        
                        if "N° DE PARTE" in df.columns:
                            df = df.dropna(subset=["N° DE PARTE"])
                        
                        # Filtro especial LIPU
                        if almacen == "ALM. LIPU" and "CONSIGNA" in df.columns:
                            df = df[df["CONSIGNA"].astype(str).str.contains("LIPU", case=False, na=False)]
                        
                        df = df.rename(columns={"N° DE PARTE": "NP", "DESCRIPCION": "DESCR"})
                        df["AGENCIA"] = agencia
                        df["ALMACEN"] = almacen

                        columnas_existentes = [col for col in columnas_finales if col in df.columns]
                        for col_faltante in set(columnas_finales) - set(columnas_existentes):
                            df[col_faltante] = None

                        df_limpio = df[columnas_finales].copy()
                        dataframes_limpios.append(df_limpio)
                    else:
                        errores.append(f"❌ Acceso Denegado a '{almacen}': El Robot no tiene permisos. (Dale compartir como Lector al correo del Robot).")

                except Exception as e:
                    errores.append(f"❌ Error al extraer '{almacen}': {str(e)}")
                
                barra.progress((i + 1) / total_links)

            for err in errores:
                st.error(err)

            # 4. Consolidar Base Madre
            if dataframes_limpios:
                df_master = pd.concat(dataframes_limpios, ignore_index=True)
                df_master["NP"] = df_master["NP"].astype(str).str.strip()
                df_master = df_master[df_master["NP"] != ""]

                df_master = df_master.merge(df_cat, on='NP', how='left')
                df_master['DESCR'] = df_master['DESCR_OFICIAL'].fillna("X DEFINIR")
                df_master = df_master.drop(columns=['DESCR_OFICIAL'])

                # BLINDAJES
                columnas_num = ["CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "DIAS ENTREGA"]
                for col in columnas_num:
                    df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)
                    df_master[col] = df_master[col].astype(int)

                df_master["HORA SOLICITUD"] = df_master["HORA SOLICITUD"].apply(normalizar_hora)
                df_master["HORA ENTREGA"] = df_master["HORA ENTREGA"].apply(normalizar_hora)

                for col_fecha in ["FECHA SOLICITUD", "FECHA ENTREGA"]:
                    df_master[col_fecha] = pd.to_datetime(df_master[col_fecha], format="%d/%m/%Y", errors='coerce')
                    if col_fecha == "FECHA ENTREGA":
                        df_master[col_fecha] = df_master[col_fecha].fillna(df_master["FECHA SOLICITUD"])
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

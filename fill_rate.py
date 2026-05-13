import streamlit as st
import pandas as pd
import io

def normalizar_hora(val):
    """Convierte cualquier formato de celda de Excel de hora a un string limpio 'HH:MM:SS' 
    o devuelve '00:00:00' si hay un error o está vacío."""
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
        return "00:00:00"
    
    # Si ya es un objeto time de datetime
    if isinstance(val, (pd.Timestamp, datetime.time if 'datetime' in globals() else type(None))):
        try:
            return val.strftime('%H:%M:%S')
        except:
            pass

    val_str = str(val).strip()
    
    # Si viene con formato de fecha y hora completo (ej. '2023-10-25 14:30:00')
    if " " in val_str:
        val_str = val_str.split(" ")[1]
        
    # Limpieza de decimales flotantes residuales si Excel lo exportó mal
    if "." in val_str and ":" not in val_str:
        try:
            # Convertir fracción de día a hora de Excel
            total_seconds = int(float(val_str) * 86400)
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        except:
            return "00:00:00"
            
    # Si tiene formato estándar HH:MM o HH:MM:SS
    partes = val_str.split(":")
    if len(partes) >= 2:
        try:
            h = int(float(partes[0]))
            m = int(float(partes[1]))
            s = int(float(partes[2])) if len(partes) > 2 else 0
            return f"{h:02d}:{m:02d}:{s:02d}"
        except:
            return "00:00:00"
            
    return "00:00:00"

def show():
    st.title("📊 Generador Base Madre: Fill Rate")
    st.markdown("Homogeneiza las columnas de tus consignas, estandariza descripciones y genera la base final para Power BI.")

    # ==========================================
    # 1. CARGA DEL CATÁLOGO MAESTRO
    # ==========================================
    st.markdown("### 📚 1. Catálogo Maestro de Refacciones")
    st.info("Sube el archivo que contiene la descripción oficial de los números de parte. (Debe tener las columnas 'NP' y 'DESCRIPCION')")
    archivo_catalogo = st.file_uploader("Cargar Catálogo Maestro", type=["xlsx", "xls", "csv"], key="cat_maestro")

    st.markdown("---")
    st.markdown("### 🗂️ 2. Carga de Archivos Operativos (Bitácoras)")

    consignas_cuautitlan = ["ALM. LIPU", "ALM. MYM", "ALM. UTEP", "ALM. UTEP SAN LUIS", "IDEALEASE"]
    consignas_tultitlan = ["ALMACEN AFN", "BISONTE SLP", "BISONTE TEPOTZOTLAN", "CULVERT", "TDR", "TUMSA"]

    columnas_finales = [
        "AGENCIA", "ALMACEN", "NP", "DESCR", "FECHA SOLICITUD (DIA/MES/AÑO)", 
        "HORA SOLICITUD", "FECHA ENTREGA (DIA/MES/AÑO)", "HORA ENTREGA (Formato 24 hrs)", 
        "DIAS ENTREGA", "CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "ESTATUS"
    ]

    archivos_subidos = []

    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("📍 CUAUTITLÁN")
        for almacen in consignas_cuautitlan:
            archivo = st.file_uploader(f"Cargar: {almacen}", type=["xlsx", "xls"], key=almacen)
            if archivo:
                archivos_subidos.append({"Agencia": "CUAUTITLAN", "Almacen": almacen, "Archivo": archivo})

    with c2:
        st.subheader("📍 TULTITLÁN")
        for almacen in consignas_tultitlan:
            archivo = st.file_uploader(f"Cargar: {almacen}", type=["xlsx", "xls"], key=almacen)
            if archivo:
                archivos_subidos.append({"Agencia": "TULTITLAN", "Almacen": almacen, "Archivo": archivo})

    # ==========================================
    # PROCESAMIENTO DE DATOS
    # ==========================================
    st.markdown("---")
    if st.button("🚀 Procesar y Generar Base Madre", type="primary"):
        if not archivos_subidos:
            st.warning("⚠️ Sube al menos un archivo operativo para procesar.")
            return
            
        if not archivo_catalogo:
            st.warning("⚠️ Por favor sube primero el Catálogo Maestro de Refacciones en el paso 1.")
            return

        try:
            if archivo_catalogo.name.endswith('.csv'):
                df_cat = pd.read_csv(archivo_catalogo)
            else:
                df_cat = pd.read_excel(archivo_catalogo)
            
            df_cat.columns = [col.strip().upper() for col in df_cat.columns]
            col_descr_cat = "DESCRIPCION" if "DESCRIPCION" in df_cat.columns else "DESCRIPCIÓN"
            
            if "NP" not in df_cat.columns or col_descr_cat not in df_cat.columns:
                st.error(f"❌ El catálogo maestro debe tener exactamente las columnas 'NP' y '{col_descr_cat}'.")
                return
                
            df_cat = df_cat[['NP', col_descr_cat]].drop_duplicates(subset=['NP'])
            # Asegurar que el NP del catálogo sea string y sin espacios truncos
            df_cat['NP'] = df_cat['NP'].astype(str).str.strip()
            df_cat = df_cat.rename(columns={col_descr_cat: "DESCR_OFICIAL"})
            
        except Exception as e:
            st.error(f"❌ Error al procesar el Catálogo Maestro: {e}")
            return

        dataframes_limpios = []
        errores = []

        for item in archivos_subidos:
            agencia = item["Agencia"]
            almacen = item["Almacen"]
            file = item["Archivo"]

            try:
                df = pd.read_excel(file, header=11)
                
                if "N° DE PARTE" in df.columns:
                    df = df.dropna(subset=["N° DE PARTE"])
                
                if almacen == "ALM. LIPU" and "CONSIGNA" in df.columns:
                    df = df[df["CONSIGNA"].astype(str).str.contains("LIPU", case=False, na=False)]
                
                # Mapeos dinámicos
                if almacen == "IDEALEASE":
                    mapeo_idealease = {
                        "N° DE PARTE": "NP", "DESCRIPCIÓN": "DESCR",
                        "FECHA SOLICITUD (DE PEDIDO)": "FECHA SOLICITUD (DIA/MES/AÑO)",
                        "HORA SOLICITUD (DE PEDIDO)": "HORA SOLICITUD",
                        "FECHA ENTREGA (DE PEDIDO)": "FECHA ENTREGA (DIA/MES/AÑO)",
                        "HORA ENTREGA (DE PEDIDO)": "HORA ENTREGA (Formato 24 hrs)",
                        "DIAS ENTREGA": "DIAS ENTREGA", "CANTIDAD SOLICITADA": "CANTIDAD SOLICITADA",
                        "CANTIDAD ENTREGADA": "CANTIDAD ENTREGADA", "ESTATUS": "ESTATUS"
                    }
                    df = df.rename(columns=mapeo_idealease)
                else:
                    mapeo_estandar = {
                        "N° DE PARTE": "NP", "DESCRIPCION": "DESCR", "DESCRIPCIÓN": "DESCR", 
                        "FECHA SOLICITUD (DIA/MES/AÑO)": "FECHA SOLICITUD (DIA/MES/AÑO)",
                        "FECHA SOLICITUD": "FECHA SOLICITUD (DIA/MES/AÑO)", 
                        "HORA SOLICITUD": "HORA SOLICITUD",
                        "FECHA ENTREGA (DIA/MES/AÑO)": "FECHA ENTREGA (DIA/MES/AÑO)",
                        "FECHA ENTREGA": "FECHA ENTREGA (DIA/MES/AÑO)", 
                        "HORA ENTREGA (Formato 24 hrs)": "HORA ENTREGA (Formato 24 hrs)",
                        "HORA ENTREGA": "HORA ENTREGA (Formato 24 hrs)", 
                        "DIAS ENTREGA": "DIAS ENTREGA", "CANTIDAD SOLICITADA": "CANTIDAD SOLICITADA",
                        "CANTIDAD ENTREGADA": "CANTIDAD ENTREGADA", "ESTATUS": "ESTATUS"
                    }
                    df = df.rename(columns=mapeo_estandar)

                df["AGENCIA"] = agencia
                df["ALMACEN"] = almacen

                columnas_existentes = [col for col in columnas_finales if col in df.columns]
                for col_faltante in set(columnas_finales) - set(columnas_existentes):
                    df[col_faltante] = None

                df_limpio = df[columnas_finales].copy()
                dataframes_limpios.append(df_limpio)

            except Exception as e:
                errores.append(f"❌ Error al procesar {almacen}: {str(e)}")

        for err in errores:
            st.error(err)

        if dataframes_limpios:
            df_master = pd.concat(dataframes_limpios, ignore_index=True)
            df_master["NP"] = df_master["NP"].astype(str).str.strip()
            df_master = df_master[df_master["NP"] != ""]

            # Cruce con el catálogo
            df_master = df_master.merge(df_cat, on='NP', how='left')
            df_master['DESCR'] = df_master['DESCR_OFICIAL'].fillna("X DEFINIR")
            df_master = df_master.drop(columns=['DESCR_OFICIAL'])

            # ==========================================
            # 🛡️ OPERACIÓN BLINDAJE ULTRA (Cero Errores en Power BI)
            # ==========================================
            
            # 1. BLINDAJE DE CANTIDADES Y DÍAS (Fila Roja Faltante -> 0)
            columnas_num = ["CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "DIAS ENTREGA"]
            for col in columnas_num:
                # Reemplazar nulos, NaN y strings vacíos por 0 antes de cambiar el tipo de dato
                df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)
                # Forzar a INT puro (Evita decimales flotantes como 30.0 que rompen DAX)
                df_master[col] = df_master[col].astype(int)

            # 2. BLINDAJE DE HORAS (Evita errores de tipos cruzados en Power BI)
            df_master["HORA SOLICITUD"] = df_master["HORA SOLICITUD"].apply(normalizar_hora)
            df_master["HORA ENTREGA (Formato 24 hrs)"] = df_master["HORA ENTREGA (Formato 24 hrs)"].apply(normalizar_hora)

            # 3. BLINDAJE DE FECHAS (Formato Homogéneo YYYY-MM-DD para evitar el choque regional de Power BI)
            for col_fecha in ["FECHA SOLICITUD (DIA/MES/AÑO)", "FECHA ENTREGA (DIA/MES/AÑO)"]:
                df_master[col_fecha] = pd.to_datetime(df_master[col_fecha], errors='coerce')
                # Si una fecha de entrega viene nula (Backorder activo), le asignamos la de solicitud para no dejar vacíos peligrosos
                if col_fecha == "FECHA ENTREGA (DIA/MES/AÑO)":
                    df_master[col_fecha] = df_master[col_fecha].fillna(df_master["FECHA SOLICITUD (DIA/MES/AÑO)"])
                # Llenar nulos remanentes crónicos con una fecha pivote o la fecha de hoy
                df_master[col_fecha] = df_master[col_fecha].fillna(pd.Timestamp('2026-01-01'))
                # Formatear como string estándar ISO (Power BI lo autodetecta sin errores de idioma)
                df_master[col_fecha] = df_master[col_fecha].dt.strftime('%Y-%m-%d')

            # ==========================================

            st.success(f"✅ ¡Base Madre generada, blindada contra Nulos y Normalizada! (Total: {len(df_master)} registros)")
            st.dataframe(df_master.head(15))

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_master.to_excel(writer, index=False, sheet_name='Fact_Fill_Rate')
            
            st.download_button(
                label="📥 Descargar Base Madre Normalizada (Excel)",
                data=buffer.getvalue(),
                file_name="BASE_MADRE_FILL_RATE.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

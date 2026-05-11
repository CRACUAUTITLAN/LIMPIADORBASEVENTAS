import streamlit as st
import pandas as pd
import io

def show():
    st.title("📊 Generador Base Madre: Fill Rate")
    st.markdown("Homogeneiza las columnas de tus consignas, estandariza descripciones y genera la base final para Power BI.")

    # ==========================================
    # 1. CARGA DEL CATÁLOGO MAESTRO (NUEVO)
    # ==========================================
    st.markdown("### 📚 1. Catálogo Maestro de Refacciones")
    st.info("Sube el archivo que contiene la descripción oficial de los números de parte. (Debe tener las columnas 'NP' y 'DESCRIPCION')")
    archivo_catalogo = st.file_uploader("Cargar Catálogo Maestro", type=["xlsx", "xls", "csv"], key="cat_maestro")

    st.markdown("---")
    st.markdown("### 🗂️ 2. Carga de Archivos Operativos (Bitácoras)")

    # Diccionarios de Agencias y Almacenes (Consignas)
    consignas_cuautitlan = ["ALM. LIPU", "ALM. MYM", "ALM. UTEP", "ALM. UTEP SAN LUIS", "IDEALEASE"]
    consignas_tultitlan = ["ALMACEN AFN", "BISONTE SLP", "BISONTE TEPOTZOTLAN", "CULVERT", "TDR", "TUMSA"]

    # Orden final estricto para Power BI
    columnas_finales = [
        "AGENCIA", "ALMACEN", "NP", "DESCR", "FECHA SOLICITUD (DIA/MES/AÑO)", 
        "HORA SOLICITUD", "FECHA ENTREGA (DIA/MES/AÑO)", "HORA ENTREGA (Formato 24 hrs)", 
        "DIAS ENTREGA", "CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "ESTATUS"
    ]

    archivos_subidos = []

    # Interfaz de subida dividida por Agencia
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

        # 1. Leer y preparar el Catálogo Maestro
        try:
            if archivo_catalogo.name.endswith('.csv'):
                df_cat = pd.read_csv(archivo_catalogo)
            else:
                df_cat = pd.read_excel(archivo_catalogo)
            
            # Limpiar nombres de columnas del catálogo para evitar errores de espacios
            df_cat.columns = [col.strip().upper() for col in df_cat.columns]
            
            # Detectar la columna de descripción (por si le pusieron acento o no)
            col_descr_cat = "DESCRIPCION" if "DESCRIPCION" in df_cat.columns else "DESCRIPCIÓN"
            
            if "NP" not in df_cat.columns or col_descr_cat not in df_cat.columns:
                st.error(f"❌ El catálogo maestro debe tener exactamente las columnas 'NP' y '{col_descr_cat}'.")
                return
                
            # Dejar solo NP y Descripcion oficial, y quitar duplicados por si acaso
            df_cat = df_cat[['NP', col_descr_cat]].drop_duplicates(subset=['NP'])
            # Renombrar para el cruce
            df_cat = df_cat.rename(columns={col_descr_cat: "DESCR_OFICIAL"})
            
        except Exception as e:
            st.error(f"❌ Error al procesar el Catálogo Maestro: {e}")
            return

        # 2. Procesar los archivos operativos
        dataframes_limpios = []
        errores = []

        for item in archivos_subidos:
            agencia = item["Agencia"]
            almacen = item["Almacen"]
            file = item["Archivo"]

            try:
                # Leer el excel omitiendo las filas azules de arriba (Fila 12 es el header)
                df = pd.read_excel(file, header=11)
                
                if "N° DE PARTE" in df.columns:
                    df = df.dropna(subset=["N° DE PARTE"])
                
                # Mapeo de columnas según el tipo de almacén
                if almacen == "IDEALEASE":
                    mapeo_idealease = {
                        "N° DE PARTE": "NP",
                        "DESCRIPCIÓN": "DESCR",
                        "FECHA SOLICITUD (DE PEDIDO)": "FECHA SOLICITUD (DIA/MES/AÑO)",
                        "HORA SOLICITUD (DE PEDIDO)": "HORA SOLICITUD",
                        "FECHA ENTREGA (DE PEDIDO)": "FECHA ENTREGA (DIA/MES/AÑO)",
                        "HORA ENTREGA (DE PEDIDO)": "HORA ENTREGA (Formato 24 hrs)",
                        "DIAS ENTREGA": "DIAS ENTREGA",
                        "CANTIDAD SOLICITADA": "CANTIDAD SOLICITADA",
                        "CANTIDAD ENTREGADA": "CANTIDAD ENTREGADA",
                        "ESTATUS": "ESTATUS"
                    }
                    df = df.rename(columns=mapeo_idealease)
                else:
                    mapeo_estandar = {
                        "N° DE PARTE": "NP",
                        "DESCRIPCION": "DESCR",
                        "DESCRIPCIÓN": "DESCR", 
                        "FECHA SOLICITUD (DIA/MES/AÑO)": "FECHA SOLICITUD (DIA/MES/AÑO)",
                        "FECHA SOLICITUD": "FECHA SOLICITUD (DIA/MES/AÑO)", 
                        "HORA SOLICITUD": "HORA SOLICITUD",
                        "FECHA ENTREGA (DIA/MES/AÑO)": "FECHA ENTREGA (DIA/MES/AÑO)",
                        "FECHA ENTREGA": "FECHA ENTREGA (DIA/MES/AÑO)", 
                        "HORA ENTREGA (Formato 24 hrs)": "HORA ENTREGA (Formato 24 hrs)",
                        "HORA ENTREGA": "HORA ENTREGA (Formato 24 hrs)", 
                        "DIAS ENTREGA": "DIAS ENTREGA",
                        "CANTIDAD SOLICITADA": "CANTIDAD SOLICITADA",
                        "CANTIDAD ENTREGADA": "CANTIDAD ENTREGADA",
                        "ESTATUS": "ESTATUS"
                    }
                    df = df.rename(columns=mapeo_estandar)

                # Agregar columnas de dimensiones
                df["AGENCIA"] = agencia
                df["ALMACEN"] = almacen

                # Verificar columnas existentes y rellenar nulos si faltan
                columnas_existentes = [col for col in columnas_finales if col in df.columns]
                for col_faltante in set(columnas_finales) - set(columnas_existentes):
                    df[col_faltante] = None

                # Seleccionar estrictamente el orden de columnas
                df_limpio = df[columnas_finales]
                dataframes_limpios.append(df_limpio)

            except Exception as e:
                errores.append(f"❌ Error al procesar {almacen}: {str(e)}")

        for err in errores:
            st.error(err)

        # 3. Consolidar Base y Normalizar Descripciones
        if dataframes_limpios:
            df_master = pd.concat(dataframes_limpios, ignore_index=True)
            df_master = df_master.dropna(subset=["NP"])

            # --- LA MAGIA: NORMALIZACIÓN CON EL CATÁLOGO ---
            # Hacemos un cruce (Left Join) con el catálogo maestro usando el NP
            df_master = df_master.merge(df_cat, on='NP', how='left')
            
            # Reemplazamos la columna DESCR por la DESCR_OFICIAL del catálogo
            # Si un NP no estaba en el catálogo, le dejamos la descripción que ya traía el archivo (fillna)
            df_master['DESCR'] = df_master['DESCR_OFICIAL'].fillna(df_master['DESCR'])
            
            # Borramos la columna temporal de cruce para que quede la tabla impecable
            df_master = df_master.drop(columns=['DESCR_OFICIAL'])

            st.success(f"✅ ¡Base Madre generada y normalizada con éxito! (Total: {len(df_master)} registros)")
            
            # Vista previa
            st.dataframe(df_master.head(15))

            # Preparar descarga Excel
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

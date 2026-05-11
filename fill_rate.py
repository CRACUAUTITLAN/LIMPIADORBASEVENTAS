import streamlit as st
import pandas as pd
import io

def show():
    st.title("📊 Generador Base Madre: Fill Rate")
    st.markdown("Sube los archivos descargados de Google Sheets. El sistema homogeneizará las columnas, agregará Agencia/Almacén y generará la base final para Power BI.")

    # Diccionarios de Agencias y Almacenes (Consignas)
    consignas_cuautitlan = ["ALM. LIPU", "ALM. MYM", "ALM. UTEP", "ALM. UTEP SAN LUIS", "IDEALEASE"]
    consignas_tultitlan = ["ALMACEN AFN", "BISONTE SLP", "BISONTE TEPOTZOTLAN", "CULVERT", "TDR", "TUMSA"]

    # Definimos el orden final estricto que queremos para Power BI
    columnas_finales = [
        "AGENCIA", "ALMACEN", "NP", "DESCR", "FECHA SOLICITUD (DIA/MES/AÑO)", 
        "HORA SOLICITUD", "FECHA ENTREGA (DIA/MES/AÑO)", "HORA ENTREGA (Formato 24 hrs)", 
        "DIAS ENTREGA", "CANTIDAD SOLICITADA", "CANTIDAD ENTREGADA", "ESTATUS"
    ]

    # Diccionario para guardar los archivos subidos
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

    # Botón de Procesamiento
    st.markdown("---")
    if st.button("🚀 Procesar y Generar Base Madre", type="primary"):
        if not archivos_subidos:
            st.warning("⚠️ Sube al menos un archivo para procesar.")
            return

        dataframes_limpios = []
        errores = []

        for item in archivos_subidos:
            agencia = item["Agencia"]
            almacen = item["Almacen"]
            file = item["Archivo"]

            try:
                # Leer el excel. header=11 significa que la fila 12 de Excel es el encabezado (índice 11 en Pandas)
                df = pd.read_excel(file, header=11)
                
                # Quitar filas donde el NP esté vacío (filas en blanco de Excel)
                if "N° DE PARTE" in df.columns:
                    df = df.dropna(subset=["N° DE PARTE"])
                
                # Mapeo de columnas según el tipo de almacén
                if almacen == "IDEALEASE":
                    # Lógica especial para IDEALEASE
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
                    # Lógica estándar para las demás consignas
                    mapeo_estandar = {
                        "N° DE PARTE": "NP",
                        "DESCRIPCION": "DESCR",
                        "DESCRIPCIÓN": "DESCR", # Cubrimos ambas ortografías por si acaso
                        "FECHA SOLICITUD (DIA/MES/AÑO)": "FECHA SOLICITUD (DIA/MES/AÑO)",
                        "FECHA SOLICITUD": "FECHA SOLICITUD (DIA/MES/AÑO)", # Mapeo extra de seguridad
                        "HORA SOLICITUD": "HORA SOLICITUD",
                        "FECHA ENTREGA (DIA/MES/AÑO)": "FECHA ENTREGA (DIA/MES/AÑO)",
                        "FECHA ENTREGA": "FECHA ENTREGA (DIA/MES/AÑO)", # Mapeo extra de seguridad
                        "HORA ENTREGA (Formato 24 hrs)": "HORA ENTREGA (Formato 24 hrs)",
                        "HORA ENTREGA": "HORA ENTREGA (Formato 24 hrs)", # Mapeo extra de seguridad
                        "DIAS ENTREGA": "DIAS ENTREGA",
                        "CANTIDAD SOLICITADA": "CANTIDAD SOLICITADA",
                        "CANTIDAD ENTREGADA": "CANTIDAD ENTREGADA",
                        "ESTATUS": "ESTATUS"
                    }
                    df = df.rename(columns=mapeo_estandar)

                # Agregar columnas de dimensiones
                df["AGENCIA"] = agencia
                df["ALMACEN"] = almacen

                # Verificar que existan las columnas necesarias antes de filtrar
                columnas_existentes = [col for col in columnas_finales if col in df.columns]
                
                # Rellenar con nulo las columnas que falten (por si hay algún error en el formato de origen)
                for col_faltante in set(columnas_finales) - set(columnas_existentes):
                    df[col_faltante] = None

                # Seleccionar estrictamente el orden de columnas final
                df_limpio = df[columnas_finales]
                
                dataframes_limpios.append(df_limpio)

            except Exception as e:
                errores.append(f"❌ Error al procesar {almacen}: {str(e)}")

        # Mostrar errores si los hay
        for err in errores:
            st.error(err)

        # Si hay dataframes exitosos, consolidar
        if dataframes_limpios:
            # CORRECCIÓN APLICADA AQUÍ: ignore_index=True
            df_master = pd.concat(dataframes_limpios, ignore_index=True)
            
            # Limpieza final (quitar filas basura si se colaron nulos en NP)
            df_master = df_master.dropna(subset=["NP"])

            st.success(f"✅ ¡Base Madre generada con éxito! (Total de registros procesados: {len(df_master)})")
            
            # Vista previa
            st.dataframe(df_master.head(15))

            # Preparar descarga Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_master.to_excel(writer, index=False, sheet_name='Fact_Fill_Rate')
            
            st.download_button(
                label="📥 Descargar Base Madre (Excel)",
                data=buffer.getvalue(),
                file_name="BASE_MADRE_FILL_RATE.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

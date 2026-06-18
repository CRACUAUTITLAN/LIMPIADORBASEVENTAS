import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io

def show():
    st.title("🏭 ALMACÉN ACTUAL IDEALEASE")
    st.write("Carga tu reporte diario para generar la base general limpia.")

    # Ya solo necesitamos subir un archivo, así que quitamos las dos columnas
    file_reporte = st.file_uploader("📥 Sube el Reporte Diario (Debe ser formato Excel .xlsx)", type=['xlsx'])

    if file_reporte:
        # Validación preventiva para evitar caídas del sistema con archivos CSV
        if file_reporte.name.lower().endswith('.csv'):
            st.error("🚫 El archivo seleccionado es un CSV (.csv). Por favor, sube el archivo original directamente en formato de Excel (.xlsx).")
            return

        if st.button("🚀 Procesar y Limpiar Base"):
            with st.spinner("Procesando estructura y textos..."):
                try:
                    # 1. Cargar datos con pandas
                    df = pd.read_excel(file_reporte)
                    
                    # Normalizar nombres de columnas a mayúsculas para evitar errores por variaciones de caja
                    df_cols_originales = [str(c).strip() for c in df.columns]
                    df.columns = [c.upper() for c in df_cols_originales]
                    
                    # Mapeo flexible de columnas clave
                    col_no = 'NO.' if 'NO.' in df.columns else ('NO' if 'NO' in df.columns else None)
                    col_eco = 'ECO' if 'ECO' in df.columns else None
                    col_obs = 'OBSERVACIONES' if 'OBSERVACIONES' in df.columns else None
                    col_estatus = 'ESTATUS' if 'ESTATUS' in df.columns else None
                    
                    if not col_no or not col_eco or not col_obs or not col_estatus:
                        st.error(f"🚫 Faltan columnas esenciales ('No.', 'ECO', 'OBSERVACIONES' o 'ESTATUS'). Columnas detectadas: {df_cols_originales}")
                        return

                    # 2. Filtrar filas donde 'No.' y 'ECO' contengan valores numéricos válidos
                    df['NO_NUM'] = pd.to_numeric(df[col_no], errors='coerce')
                    df['ECO_NUM'] = pd.to_numeric(df[col_eco], errors='coerce')
                    df_limpio = df.dropna(subset=['NO_NUM', 'ECO_NUM']).copy()

                    # Listado de series (VIN) con estatus fijo "EN RENTA"
                    vins_excepcion = [
                        "3MN2E82W4HD700055", "3MN2E82W7HD700115", "3MN2E82W7HD700051",
                        "3MN2E82W8HD700107", "3MN2E82W0HD600583", "3MN2E82W1HD600592"
                    ]

                    # 3. Iterar renglón por renglón para aplicar reglas de negocio
                    estatus_lista = []
                    cliente_lista = []
                    
                    # Expresión regular para ubicar fechas (ej. 26 JUL 24 o 9/DIC/23)
                    patron_fecha = r'\b\d{1,2}\s*[/]?\s*[a-zA-Z]+\s*[/]?\s*\d{2,4}\b'

                    for index, row in df_limpio.iterrows():
                        # --- LIMPIEZA DE CLIENTE (Texto previo a la fecha) ---
                        obs_texto = str(row[col_obs]) if pd.notna(row[col_obs]) else ""
                        cliente = re.split(patron_fecha, obs_texto)[0].strip()
                        cliente_lista.append(cliente)

                        # --- ASIGNACIÓN DE ESTATUS DESDE LA NUEVA COLUMNA O EXCEPCIÓN ---
                        vin_actual = str(row['VIN']).strip() if 'VIN' in df_limpio.columns else ""
                        
                        if vin_actual in vins_excepcion:
                            estatus_lista.append("EN RENTA")
                        else:
                            # Leemos directamente lo que Marianita capturó en la columna ESTATUS
                            estatus_capturado = str(row[col_estatus]).strip().upper() if pd.notna(row[col_estatus]) else "SIN DEFINIR"
                            estatus_lista.append(estatus_capturado)

                    # 4. Estructurar la Base de Datos Final limpia
                    df_final = pd.DataFrame({
                        "MODELO": df_limpio.get('MODELO', ''),
                        "VIN": df_limpio.get('VIN', ''),
                        "ECO": df_limpio.get('ECO', ''),
                        "PLACAS": df_limpio.get('PLACAS', ''),
                        "ADAPTACIÓN": df_limpio.get('ADAPTACIÓN', df_limpio.get('ADAPTACION', '')),
                        "AÑO": df_limpio.get('AÑO', df_limpio.get('ANO', '')),
                        "TIPO": df_limpio.get('TIPO', ''),
                        "CAPACIDAD": df_limpio.get('CAPACIDAD', ''),
                        "ESTATUS": estatus_lista,
                        "CLIENTE/TALLER/DISPONIBILIDAD": cliente_lista,
                        "FECHA": datetime.now().strftime("%d/%m/%Y")  # Formato DIA/MES/AÑO
                    })

                    st.success("✅ ¡Base de datos limpiada y procesada con éxito!")
                    st.dataframe(df_final)

                    # Configurar la descarga del nuevo archivo de salida en formato Excel
                    buffer = io.BytesIO()
                    # Seguimos usando openpyxl aquí solo como motor para escribir el Excel de salida
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Base_Limpia')
                    
                    st.download_button(
                        label="📥 Descargar Base Limpia",
                        data=buffer.getvalue(),
                        file_name=f"Almacen_Idealease_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"❌ Ocurrió un error inesperado al procesar el archivo: {e}")

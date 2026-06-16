import streamlit as st
import pandas as pd
import re
from datetime import datetime
import openpyxl
import io

def show():
    st.title("🏭 ALMACÉN ACTUAL IDEALEASE")
    st.write("Carga tu reporte diario y la regla de colores para generar la base general limpia.")

    col1, col2 = st.columns(2)
    with col1:
        file_reporte = st.file_uploader("📥 1. Sube el Reporte Diario (Debe ser formato Excel .xlsx)", type=['xlsx'])
    with col2:
        file_regla = st.file_uploader("🎨 2. Sube la Mini Regla de Colores (Opcional)", type=['xlsx'])

    if file_reporte:
        # Validación preventiva para evitar caídas del sistema con archivos CSV
        if file_reporte.name.lower().endswith('.csv'):
            st.error("🚫 El archivo seleccionado es un CSV (.csv). Los archivos CSV no conservan los colores de las celdas. Por favor, sube el archivo original directamente en formato de Excel (.xlsx).")
            return

        if st.button("🚀 Procesar y Limpiar Base"):
            with st.spinner("Procesando estructura, colores y textos..."):
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
                    
                    if not col_no or not col_eco or not col_obs:
                        st.error(f"🚫 No se encontraron las columnas esenciales ('No.', 'ECO' u 'OBSERVACIONES') en el reporte. Columnas detectadas: {df_cols_originales}")
                        return
                    
                    # 2. Cargar con openpyxl para extraer los colores
                    wb = openpyxl.load_workbook(file_reporte, data_only=True)
                    sheet = wb.active

                    # 3. Filtrar filas donde 'No.' y 'ECO' contengan valores numéricos válidos
                    df['NO_NUM'] = pd.to_numeric(df[col_no], errors='coerce')
                    df['ECO_NUM'] = pd.to_numeric(df[col_eco], errors='coerce')
                    df_limpio = df.dropna(subset=['NO_NUM', 'ECO_NUM']).copy()

                    # Encontrar el índice de la columna OBSERVACIONES en la hoja de cálculo de Excel (Base 1)
                    idx_obs = [c.upper() for c in df_cols_originales].index('OBSERVACIONES') + 1

                    # 4. Procesar la Mini Regla de Colores (Si se proporciona)
                    mapa_colores = {}
                    if file_regla:
                        wb_regla = openpyxl.load_workbook(file_regla, data_only=True)
                        sheet_regla = wb_regla.active
                        for row in sheet_regla.iter_rows(min_row=1, max_row=20):
                            for cell in row:
                                val = str(cell.value).strip().upper() if cell.value else ""
                                if val in ["DISPONIBLE", "EN RENTA", "TALLER"]:
                                    if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                                        mapa_colores[cell.fill.start_color.rgb] = val

                    # Listado de series (VIN) con estatus fijo "EN RENTA"
                    vins_excepcion = [
                        "3MN2E82W4HD700055", "3MN2E82W7HD700115", "3MN2E82W7HD700051",
                        "3MN2E82W8HD700107", "3MN2E82W0HD600583", "3MN2E82W1HD600592"
                    ]

                    # 5. Iterar renglón por renglón para aplicar reglas de negocio
                    estatus_lista = []
                    cliente_lista = []
                    
                    # Expresión regular para ubicar fechas (ej. 26 JUL 24 o 9/DIC/23)
                    patron_fecha = r'\b\d{1,2}\s*[/]?\s*[a-zA-Z]+\s*[/]?\s*\d{2,4}\b'

                    for index, row in df_limpio.iterrows():
                        # Ajustar el índice de pandas a la fila exacta de Excel (encabezado en fila 1)
                        excel_row = index + 2
                        
                        # --- LIMPIEZA DE CLIENTE (Texto previo a la fecha) ---
                        obs_texto = str(row['OBSERVACIONES']) if pd.notna(row['OBSERVACIONES']) else ""
                        cliente = re.split(patron_fecha, obs_texto)[0].strip()
                        cliente_lista.append(cliente)

                        # --- ASIGNACIÓN DE ESTATUS POR COLOR O EXCEPCIÓN ---
                        vin_actual = str(row['VIN']).strip() if 'VIN' in row else ""
                        
                        if vin_actual in vins_excepcion:
                            estatus_lista.append("EN RENTA")
                        else:
                            celda = sheet.cell(row=excel_row, column=idx_obs)
                            color_rgb = celda.fill.start_color.rgb if celda.fill and celda.fill.start_color else None
                            
                            estatus_asignado = "SIN DEFINIR"
                            if color_rgb in mapa_colores:
                                estatus_asignado = mapa_colores[color_rgb]
                            elif color_rgb and isinstance(color_rgb, str):
                                color_str = color_rgb.upper()
                                # Paletas de verdes comunes (DISPONIBLE)
                                if any(x in color_str for x in ["FF00FF00", "FF92D050", "FFC6EFCE", "00B050"]):
                                    estatus_asignado = "DISPONIBLE"
                                # Paletas de amarillos comunes (EN RENTA)
                                elif any(x in color_str for x in ["FFFFFF00", "FFFFC000", "FFFFE699", "FFEBFFF0"]):
                                    estatus_asignado = "EN RENTA"
                                # Paletas de naranjas/rojos comunes (TALLER)
                                elif any(x in color_str for x in ["FFED7D31", "FFFCE4D6", "FFC00000", "FFD96B27"]):
                                    estatus_asignado = "TALLER"
                            
                            estatus_lista.append(estatus_asignado)

                    # 6. Estructurar la Base de Datos Final limpia
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
                        "FECHA": datetime.now().strftime("%d/%m/%Y")  # Formato corregido a DIA/MES/AÑO completo
                    })

                    st.success("✅ ¡Base de datos limpiada y procesada con éxito!")
                    st.dataframe(df_final)

                    # Configurar la descarga del nuevo archivo de salida en formato Excel
                    buffer = io.BytesIO()
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

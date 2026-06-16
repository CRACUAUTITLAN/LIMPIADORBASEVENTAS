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
        file_reporte = st.file_uploader("📥 1. Sube el Reporte Diario (Excel)", type=['xlsx'])
    with col2:
        file_regla = st.file_uploader("🎨 2. Sube la Mini Regla de Colores (Opcional)", type=['xlsx'])

    if file_reporte:
        if st.button("🚀 Procesar y Limpiar Base"):
            with st.spinner("Procesando estructura, colores y textos..."):
                try:
                    # 1. Cargar datos
                    df = pd.read_excel(file_reporte)
                    
                    # 2. Cargar con openpyxl para leer los fondos de color de la columna OBSERVACIONES
                    wb = openpyxl.load_workbook(file_reporte, data_only=True)
                    sheet = wb.active

                    # 3. Filtrar filas donde 'No.' y 'ECO' tengan números (Ignoramos vacíos o textos)
                    df['No_num'] = pd.to_numeric(df['No.'], errors='coerce')
                    df['ECO_num'] = pd.to_numeric(df['ECO'], errors='coerce')
                    df_limpio = df.dropna(subset=['No_num', 'ECO_num']).copy()

                    # Encontrar en qué número de columna de Excel está 'OBSERVACIONES'
                    try:
                        idx_obs = df.columns.get_loc("OBSERVACIONES") + 1
                    except KeyError:
                        st.error("No se encontró la columna 'OBSERVACIONES' en el reporte.")
                        return

                    # 4. Procesar la Mini Regla de Colores (Si se subió)
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

                    # Excepciones de VIN que siempre deben ser "EN RENTA"
                    vins_excepcion = [
                        "3MN2E82W4HD700055", "3MN2E82W7HD700115", "3MN2E82W7HD700051",
                        "3MN2E82W8HD700107", "3MN2E82W0HD600583", "3MN2E82W1HD600592"
                    ]

                    # 5. Iterar filas para armar columnas calculadas
                    estatus_lista = []
                    cliente_lista = []
                    
                    # Patrón Regex: Buscar Número(s) seguido de Texto (Mes) y Número(s)
                    # Detecta "26 JUL 24", "9/DIC/23", "20 ABR 26"
                    patron_fecha = r'\b\d{1,2}\s*[/]?\s*[a-zA-Z]+\s*[/]?\s*\d{2,4}\b'

                    for index, row in df_limpio.iterrows():
                        # Calcular la fila exacta en Excel (index de pandas + 2 por el encabezado y base 0)
                        excel_row = index + 2
                        
                        # --- CÁLCULO DE CLIENTE/TALLER ---
                        obs_texto = str(row['OBSERVACIONES']) if pd.notna(row['OBSERVACIONES']) else ""
                        cliente = re.split(patron_fecha, obs_texto)[0].strip()
                        cliente_lista.append(cliente)

                        # --- CÁLCULO DE ESTATUS ---
                        vin_actual = str(row['VIN']).strip()
                        if vin_actual in vins_excepcion:
                            estatus_lista.append("EN RENTA")
                        else:
                            celda = sheet.cell(row=excel_row, column=idx_obs)
                            color_rgb = celda.fill.start_color.rgb if celda.fill and celda.fill.start_color else None
                            
                            estatus_asignado = "SIN DEFINIR"
                            if color_rgb in mapa_colores:
                                estatus_asignado = mapa_colores[color_rgb]
                            elif color_rgb and type(color_rgb) == str:
                                # Fallback automático de colores HEX estándar de Excel si no se sube regla
                                color_str = color_rgb.upper()
                                if "FF00" in color_str or "B050" in color_str or "92D050" in color_str:
                                    estatus_asignado = "DISPONIBLE" # Verdes
                                elif "FFFF" in color_str and "C000" not in color_str:
                                    estatus_asignado = "EN RENTA" # Amarillos
                                elif "C000" in color_str or "FFC" in color_str:
                                    estatus_asignado = "TALLER" # Naranjas
                            
                            estatus_lista.append(estatus_asignado)

                    # 6. Construir DataFrame final
                    df_final = pd.DataFrame({
                        "MODELO": df_limpio['Modelo'],
                        "VIN": df_limpio['VIN'],
                        "ECO": df_limpio['ECO'],
                        "PLACAS": df_limpio['Placas'],
                        "ADAPTACIÓN": df_limpio['Adaptación'],
                        "AÑO": df_limpio['Año'],
                        "TIPO": df_limpio['Tipo'],
                        "CAPACIDAD": df_limpio['CAPACIDAD'],
                        "ESTATUS": estatus_lista,
                        "CLIENTE/TALLER/DISPONIBILIDAD": cliente_lista,
                        "FECHA": datetime.now().strftime("%m/%Y") 
                    })

                    st.success("✅ ¡Base de datos limpiada y procesada con éxito!")
                    st.dataframe(df_final)

                    # Botón para descargar el Excel limpio
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
                    st.error(f"Ocurrió un error al procesar el archivo: {e}")

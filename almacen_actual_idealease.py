import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
import io

def show():
    st.title("🏭 ALMACÉN ACTUAL IDEALEASE")
    st.write("Genera tu base limpia del día y agrégala a tu histórico acumulado.")

    # Selector de Fecha (Calendario)
    fecha_seleccionada = st.date_input("📅 Selecciona la fecha del reporte", datetime.now().date())

    col1, col2 = st.columns(2)
    with col1:
        file_reporte = st.file_uploader("📥 1. Sube el Reporte Diario (.xlsx)", type=['xlsx'])
    with col2:
        file_historico = st.file_uploader("📚 2. Sube la Base Histórica / Acumulado (Opcional)", type=['xlsx'])

    if file_reporte:
        if file_reporte.name.lower().endswith('.csv'):
            st.error("🚫 Por favor, sube el reporte diario en formato Excel (.xlsx), no en CSV.")
            return

        if st.button("🚀 Procesar y Generar Base"):
            with st.spinner("Procesando información, normalizando datos y actualizando histórico..."):
                try:
                    # 1. Cargar reporte diario
                    df = pd.read_excel(file_reporte)
                    
                    # Normalizar columnas
                    df_cols_originales = [str(c).strip() for c in df.columns]
                    df.columns = [c.upper() for c in df_cols_originales]
                    
                    col_no = 'NO.' if 'NO.' in df.columns else ('NO' if 'NO' in df.columns else None)
                    col_eco = 'ECO' if 'ECO' in df.columns else None
                    col_obs = 'OBSERVACIONES' if 'OBSERVACIONES' in df.columns else None
                    col_estatus = 'ESTATUS' if 'ESTATUS' in df.columns else None
                    
                    if not col_no or not col_eco or not col_obs or not col_estatus:
                        st.error(f"🚫 Faltan columnas esenciales. Asegúrate de tener 'No.', 'ECO', 'OBSERVACIONES' y 'ESTATUS'.")
                        return

                    # 2. Filtrar filas válidas
                    df['NO_NUM'] = pd.to_numeric(df[col_no], errors='coerce')
                    df['ECO_NUM'] = pd.to_numeric(df[col_eco], errors='coerce')
                    df_limpio = df.dropna(subset=['NO_NUM', 'ECO_NUM']).copy()

                    vins_excepcion = [
                        "3MN2E82W4HD700055", "3MN2E82W7HD700115", "3MN2E82W7HD700051",
                        "3MN2E82W8HD700107", "3MN2E82W0HD600583", "3MN2E82W1HD600592"
                    ]

                    # 3. Procesar datos diarios
                    estatus_lista = []
                    cliente_lista = []
                    patron_fecha = r'\b\d{1,2}\s*[/]?\s*[a-zA-Z]+\s*[/]?\s*\d{2,4}\b'

                    for index, row in df_limpio.iterrows():
                        # Cliente
                        obs_texto = str(row[col_obs]) if pd.notna(row[col_obs]) else ""
                        cliente = re.split(patron_fecha, obs_texto)[0].strip()
                        cliente_lista.append(cliente)

                        # Estatus inteligente
                        vin_actual = str(row['VIN']).strip() if 'VIN' in df_limpio.columns else ""
                        
                        if vin_actual in vins_excepcion:
                            estatus_lista.append("EN RENTA")
                        else:
                            estatus_raw = str(row[col_estatus]).strip().upper() if pd.notna(row[col_estatus]) else ""
                            
                            if "REEMPLAZO" in estatus_raw:
                                estatus_lista.append("EN REEMPLAZO RENTA")
                            elif "RENTA" in estatus_raw:
                                estatus_lista.append("EN RENTA")
                            elif "TALLER" in estatus_raw:
                                estatus_lista.append("EN TALLER")
                            elif "DISPONIBLE" in estatus_raw:
                                estatus_lista.append("DISPONIBLE")
                            else:
                                estatus_lista.append("SIN DEFINIR")

                    # 4. Crear DataFrame del día
                    df_dia = pd.DataFrame({
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
                        "FECHA": fecha_seleccionada.strftime("%d/%m/%Y")
                    })

                    # 5. Lógica del Acumulado (Concatenar si hay histórico)
                    if file_historico:
                        df_historico = pd.read_excel(file_historico)
                        df_final = pd.concat([df_historico, df_dia], ignore_index=True)
                        mensaje_exito = "✅ ¡Base limpiada, normalizada y agregada al histórico correctamente!"
                    else:
                        df_final = df_dia
                        mensaje_exito = "✅ ¡Base diaria limpiada y normalizada con éxito! (No se adjuntó histórico)"

                    # --- 6. NORMALIZACIÓN DE TIPOS DE DATO ---
                    # Limpiamos cadenas vacías o puros espacios convirtiéndolos a NaN reales temporalmente
                    df_final = df_final.replace(r'^\s*$', np.nan, regex=True)

                    # Forzamos AÑO a ser estrictamente numérico entero (Int64 soporta valores nulos sin romper el tipo entero)
                    df_final['AÑO'] = pd.to_numeric(df_final['AÑO'], errors='coerce').astype('Int64')

                    # Para el resto de las columnas, aseguramos que sean texto y llenamos vacíos con "SIN DATO"
                    cols_texto = ["MODELO", "VIN", "ECO", "PLACAS", "ADAPTACIÓN", "TIPO", "CAPACIDAD", "ESTATUS", "CLIENTE/TALLER/DISPONIBILIDAD", "FECHA"]
                    for col in cols_texto:
                        if col in df_final.columns:
                            df_final[col] = df_final[col].fillna("SIN DATO").astype(str)

                    st.success(mensaje_exito)
                    st.dataframe(df_final)

                    # 7. Descarga del archivo final
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Base_Limpia')
                    
                    nombre_archivo = f"Acumulado_Idealease_{fecha_seleccionada.strftime('%Y%m%d')}.xlsx" if file_historico else f"Diario_Idealease_{fecha_seleccionada.strftime('%Y%m%d')}.xlsx"
                    
                    st.download_button(
                        label="📥 Descargar Base Final",
                        data=buffer.getvalue(),
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"❌ Ocurrió un error inesperado al procesar: {e}")

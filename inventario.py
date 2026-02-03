import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

def show():
    st.header("🏥 Limpiador de Inventarios (ETL)")
    st.markdown("Prepara los archivos crudos para la carpeta `DATA_INVENTARIO` del Power BI.")
    st.markdown("---")

    # ======================================================
    # 1. CONFIGURACIÓN FIJA (Tus Sucursales y Almacenes)
    # ======================================================
    ALMACENES_POR_SUCURSAL = {
        "CUAUTITLAN": [
            "ALM. GENERAL", "ALM. LIPU", "ALM. MYM", "Alm. Utep", "ALM. FAST FOOD", "ALM. BOÑAR"
        ],
        "TULTITLAN": [
            "ALM. GENERAL", "TUMSA", "BISONTE TEPOTZOTLAN", "TEISA", "CULVERT", "TDR", "ALMACEN AFN", "ZONTE"
        ],
        "BAJIO": [
            "ALM. UTEP SAN LUIS", "BISONTE SLP"
        ]
    }

    # ======================================================
    # 2. FUNCIONES
    # ======================================================
    def cargar_catalogo_blindado(uploaded_file):
        if uploaded_file is None: return None
        try:
            df = pd.read_excel(uploaded_file)
            col_name = df.columns[0]
            if 'Numero de Parte' in df.columns: col_name = 'Numero de Parte'
            if 'NP' in df.columns: col_name = 'NP'
            
            df.rename(columns={col_name: 'NP'}, inplace=True)
            df['NP'] = df['NP'].astype(str).str.strip()
            df = df.drop_duplicates(subset=['NP'], keep='first')
            
            cols_req = ['NP']
            if 'LINEA' in df.columns: cols_req.append('LINEA')
            if 'CATEGORIA' in df.columns: cols_req.append('CATEGORIA')
            
            return df[cols_req]
        except Exception as e:
            st.error(f"Error en catálogo: {e}")
            return None

    # ======================================================
    # 3. INTERFAZ DE CARGA
    # ======================================================
    
    file_cat = st.file_uploader("📂 Cargar Catálogo de Partes (.xlsx) (Opcional)", type=["xlsx"], key="inv_cat")

    inventarios_cargados = []

    st.markdown("### 📥 Carga de Inventarios por Agencia")
    st.info("ℹ️ Nota: Carga los archivos en todas las pestañas necesarias y al final presiona el botón Generar.")
    
    tabs = st.tabs(list(ALMACENES_POR_SUCURSAL.keys()))

    for i, sucursal in enumerate(ALMACENES_POR_SUCURSAL.keys()):
        with tabs[i]:
            st.markdown(f"**Almacenes de {sucursal}**")
            lista_almacenes = ALMACENES_POR_SUCURSAL[sucursal]
            col_a, col_b = st.columns(2)
            
            for idx, almacen in enumerate(lista_almacenes):
                donde_pintar = col_a if idx % 2 == 0 else col_b
                with donde_pintar:
                    f_inv = st.file_uploader(f"📦 {almacen}", type=["xls", "xlsx"], key=f"inv_{sucursal}_{almacen}")
                    
                    if f_inv:
                        try:
                            # 1. LECTURA RAW (Sin encabezados)
                            df_r = pd.read_excel(f_inv, header=None)
                            
                            # 2. SELECCIÓN QUIRÚRGICA DE COLUMNAS (Indices base 0)
                            # A=0, B=1, E=4, I=8, J=9, K=10, L=11, N=13
                            df_c = pd.DataFrame()
                            
                            # Datos Básicos
                            df_c['NP'] = df_r.iloc[:, 0].astype(str).str.strip()
                            df_c['DESCRIPCION'] = df_r.iloc[:, 1].astype(str).str.strip()
                            
                            # Métricas
                            df_c['COSTO_UNITARIO'] = pd.to_numeric(df_r.iloc[:, 4], errors='coerce').fillna(0)
                            df_c['EXISTENCIA'] = pd.to_numeric(df_r.iloc[:, 8], errors='coerce').fillna(0)
                            df_c['VALOR TOTAL'] = pd.to_numeric(df_r.iloc[:, 13], errors='coerce').fillna(0) # Columna N
                            
                            # Fechas (La clave del filtro)
                            df_c['FEC_INGRESO'] = pd.to_datetime(df_r.iloc[:, 9], errors='coerce')
                            df_c['FEC_ULT_COMPRA'] = pd.to_datetime(df_r.iloc[:, 10], errors='coerce')
                            df_c['FEC_ULT_VENTA_INV'] = pd.to_datetime(df_r.iloc[:, 11], errors='coerce')
                            
                            # 3. FILTRO MAESTRO: Si las 3 fechas son Nulas, la fila es basura
                            mask_validos = ~(
                                df_c['FEC_INGRESO'].isna() & 
                                df_c['FEC_ULT_COMPRA'].isna() & 
                                df_c['FEC_ULT_VENTA_INV'].isna()
                            )
                            df_c = df_c[mask_validos].copy()
                            
                            # Formateamos fechas a solo fecha (sin hora) después del filtro
                            for col_fecha in ['FEC_INGRESO', 'FEC_ULT_COMPRA', 'FEC_ULT_VENTA_INV']:
                                df_c[col_fecha] = df_c[col_fecha].dt.date

                            # Etiquetas
                            df_c['SUCURSAL'] = sucursal
                            df_c['ALMACEN'] = almacen
                            
                            if not df_c.empty:
                                inventarios_cargados.append(df_c)
                                st.success(f"✅ {len(df_c)} registros válidos (Filtrados por fecha).")
                            else:
                                st.warning("⚠️ Archivo leído, pero no se encontraron filas con fechas válidas.")
                                
                        except Exception as e:
                            st.error(f"❌ Error en {almacen}. Verifica que el Excel tenga datos en las columnas A,B,E,I,J,K,L,N.")

    # ======================================================
    # 4. CONSOLIDACIÓN Y EXPORTACIÓN
    # ======================================================
    st.markdown("---")
    if st.button("💾 GENERAR ARCHIVO LIMPIO PARA POWER BI", type="primary", use_container_width=True):
        if inventarios_cargados:
            with st.spinner("Unificando formatos y cruzando catálogo..."):
                # 1. Unir todo
                df_final = pd.concat(inventarios_cargados, ignore_index=True)
                
                # 2. Pegar Catálogo (si se subió)
                if file_cat:
                    df_catalogo = cargar_catalogo_blindado(file_cat)
                    if df_catalogo is not None:
                        df_final = pd.merge(df_final, df_catalogo, on='NP', how='left')
                        
                # 3. Rellenar vacíos de catálogo
                if 'CATEGORIA' not in df_final.columns: df_final['CATEGORIA'] = "X DEFINIR"
                if 'LINEA' not in df_final.columns: df_final['LINEA'] = "X DEFINIR"
                
                df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna("X DEFINIR")
                df_final['LINEA'] = df_final['LINEA'].fillna("X DEFINIR")
                
                # 4. ORDEN FINAL DE COLUMNAS
                cols_orden = [
                    "SUCURSAL", 
                    "ALMACEN", 
                    "NP", 
                    "DESCRIPCION", 
                    "LINEA", 
                    "CATEGORIA", 
                    "EXISTENCIA", 
                    "COSTO_UNITARIO", 
                    "VALOR TOTAL", 
                    "FEC_INGRESO", 
                    "FEC_ULT_COMPRA", 
                    "FEC_ULT_VENTA_INV"
                ]
                
                for c in cols_orden:
                    if c not in df_final.columns: df_final[c] = ""
                    
                df_export = df_final[cols_orden]
                
                # 5. Generar Excel (NOMBRE FIJO PARA REEMPLAZO EN DRIVE)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='DATA_INVENTARIO')
                
                # --- CAMBIO APLICADO: NOMBRE ESTÁTICO ---
                nombre_archivo = "INVENTARIO_CRA.xlsx"
                
                st.success(f"¡Listo! Archivo generado con {len(df_export)} registros limpios.")
                st.warning("⚠️ Guarda este archivo reemplazando el existente en: BI_REFACCIONES_MASTER > DATA RESOURCES > DATA_INVENTARIO")
                
                st.download_button(
                    label="📥 Descargar INVENTARIO_CRA.xlsx",
                    data=buffer.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("No has cargado ningún inventario válido.")

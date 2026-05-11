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
    # 1. CONFIGURACIÓN FIJA
    # ======================================================
    ALMACENES_POR_SUCURSAL = {
        "CUAUTITLAN": [
            "ALM. GENERAL", "ALM. LIPU", "ALM. MYM", "Alm. Utep", "ALM. UTEP SAN LUIS", "ALM. FAST FOOD", "ALM. BOÑAR"
        ],
        "TULTITLAN": [
            "ALM. GENERAL", "TUMSA", "BISONTE TEPOTZOTLAN", "BISONTE SLP", "TEISA", "CULVERT", "TDR", "ALMACEN AFN", "ZONTE", "ALM. ENLACES LOGISTICOS"
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
    st.info("ℹ️ Nota: Carga los archivos en todas las pestañas y al final presiona Generar.")
    
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
                            # 1. LECTURA RAW
                            df_r = pd.read_excel(f_inv, header=None)
                            df_c = pd.DataFrame()
                            
                            # Datos Básicos
                            df_c['NP'] = df_r.iloc[:, 0].astype(str).str.strip()
                            df_c['DESCRIPCION'] = df_r.iloc[:, 1].astype(str).str.strip()
                            
                            # Métricas
                            df_c['COSTO_UNITARIO'] = pd.to_numeric(df_r.iloc[:, 4], errors='coerce').fillna(0)
                            df_c['EXISTENCIA'] = pd.to_numeric(df_r.iloc[:, 8], errors='coerce').fillna(0)
                            df_c['VALOR TOTAL'] = pd.to_numeric(df_r.iloc[:, 13], errors='coerce').fillna(0)
                            
                            # --- CORRECCIÓN DE FECHAS (dayfirst=True) ---
                            # Esto es clave: le dice a Python que el primer número es el DÍA
                            df_c['FEC_INGRESO'] = pd.to_datetime(df_r.iloc[:, 9], errors='coerce', dayfirst=True)
                            df_c['FEC_ULT_COMPRA'] = pd.to_datetime(df_r.iloc[:, 10], errors='coerce', dayfirst=True)
                            df_c['FEC_ULT_VENTA_INV'] = pd.to_datetime(df_r.iloc[:, 11], errors='coerce', dayfirst=True)
                            
                            # Filtro de validez
                            mask_validos = ~(
                                df_c['FEC_INGRESO'].isna() & 
                                df_c['FEC_ULT_COMPRA'].isna() & 
                                df_c['FEC_ULT_VENTA_INV'].isna()
                            )
                            df_c = df_c[mask_validos].copy()
                            
                            # Quitamos la hora para dejar fecha limpia
                            for col_fecha in ['FEC_INGRESO', 'FEC_ULT_COMPRA', 'FEC_ULT_VENTA_INV']:
                                df_c[col_fecha] = df_c[col_fecha].dt.date

                            # Etiquetas
                            df_c['SUCURSAL'] = sucursal
                            df_c['ALMACEN'] = almacen
                            
                            if not df_c.empty:
                                inventarios_cargados.append(df_c)
                                st.success(f"✅ {len(df_c)} regs.")
                            else:
                                st.warning("⚠️ Sin datos válidos.")
                                
                        except Exception as e:
                            st.error(f"❌ Error en {almacen}: {e}")

    # ======================================================
    # 4. EXPORTACIÓN
    # ======================================================
    st.markdown("---")
    if st.button("💾 GENERAR ARCHIVO CORREGIDO", type="primary", use_container_width=True):
        if inventarios_cargados:
            with st.spinner("Procesando fechas correctamente..."):
                df_final = pd.concat(inventarios_cargados, ignore_index=True)
                
                if file_cat:
                    df_catalogo = cargar_catalogo_blindado(file_cat)
                    if df_catalogo is not None:
                        df_final = pd.merge(df_final, df_catalogo, on='NP', how='left')
                        
                if 'CATEGORIA' not in df_final.columns: df_final['CATEGORIA'] = "X DEFINIR"
                if 'LINEA' not in df_final.columns: df_final['LINEA'] = "X DEFINIR"
                df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna("X DEFINIR")
                df_final['LINEA'] = df_final['LINEA'].fillna("X DEFINIR")
                
                cols_orden = [
                    "SUCURSAL", "ALMACEN", "NP", "DESCRIPCION", "LINEA", "CATEGORIA", 
                    "EXISTENCIA", "COSTO_UNITARIO", "VALOR TOTAL", 
                    "FEC_INGRESO", "FEC_ULT_COMPRA", "FEC_ULT_VENTA_INV"
                ]
                
                for c in cols_orden:
                    if c not in df_final.columns: df_final[c] = ""
                    
                df_export = df_final[cols_orden]
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='DATA_INVENTARIO')
                    
                    # Formato visual dd/mm/yyyy en Excel para que TÚ lo veas bien
                    # Power BI leerá el dato correcto por debajo
                    workbook = writer.book
                    worksheet = writer.sheets['DATA_INVENTARIO']
                    formato_fecha = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                    worksheet.set_column('J:L', 12, formato_fecha)
                    worksheet.set_column('A:D', 15)

                nombre_archivo = "INVENTARIO_CRA.xlsx"
                st.success(f"¡Listo! Fechas corregidas (Día/Mes).")
                st.download_button(
                    label="📥 Descargar INVENTARIO_CRA.xlsx",
                    data=buffer.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("No hay datos.")

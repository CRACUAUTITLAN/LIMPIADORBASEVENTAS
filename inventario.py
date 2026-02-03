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
            "ALM. GENERAL", 
            "ALM. LIPU", 
            "ALM. MYM", 
            "Alm. Utep", 
            "ALM. FAST FOOD", 
            "ALM. BOÑAR"
        ],
        "TULTITLAN": [
            "ALM. GENERAL", 
            "TUMSA", 
            "BISONTE TEPOTZOTLAN", 
            "TEISA", 
            "CULVERT", 
            "TDR", 
            "ALMACEN AFN", 
            "ZONTE"
        ],
        "BAJIO": [
            "ALM. UTEP SAN LUIS", 
            "BISONTE SLP"
        ]
    }

    # ======================================================
    # 2. FUNCIONES
    # ======================================================
    def cargar_catalogo_blindado(uploaded_file):
        if uploaded_file is None: return None
        try:
            df = pd.read_excel(uploaded_file)
            # Buscamos la columna NP
            col_name = df.columns[0] # Asumimos la primera
            # Si tiene nombre explícito lo usamos
            if 'Numero de Parte' in df.columns: col_name = 'Numero de Parte'
            if 'NP' in df.columns: col_name = 'NP'
            
            df.rename(columns={col_name: 'NP'}, inplace=True)
            df['NP'] = df['NP'].astype(str).str.strip()
            # Quitamos duplicados para no inflar el inventario
            df = df.drop_duplicates(subset=['NP'], keep='first')
            
            # Verificamos que existan las columnas de clasificacion
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
    
    # Solo pedimos el catálogo para enriquecer la data (Línea/Familia)
    file_cat = st.file_uploader("📂 Cargar Catálogo de Partes (.xlsx) (Opcional pero recomendado)", type=["xlsx"], key="inv_cat")

    inventarios_cargados = []

    st.markdown("### 📥 Carga de Inventarios por Agencia")
    
    # Creamos las pestañas o columnas por sucursal
    tabs = st.tabs(list(ALMACENES_POR_SUCURSAL.keys()))

    for i, sucursal in enumerate(ALMACENES_POR_SUCURSAL.keys()):
        with tabs[i]:
            st.markdown(f"**Almacenes de {sucursal}**")
            lista_almacenes = ALMACENES_POR_SUCURSAL[sucursal]
            
            # Grid de carga (2 columnas para que no se vea eterno)
            col_a, col_b = st.columns(2)
            
            for idx, almacen in enumerate(lista_almacenes):
                # Alternamos columnas
                donde_pintar = col_a if idx % 2 == 0 else col_b
                
                with donde_pintar:
                    f_inv = st.file_uploader(f"📦 {almacen}", type=["xls", "xlsx"], key=f"inv_{sucursal}_{almacen}")
                    
                    if f_inv:
                        try:
                            # Lectura sin encabezados (posicional)
                            df_r = pd.read_excel(f_inv, header=None)
                            
                            df_c = pd.DataFrame()
                            # Mapeo Estricto: A(0)=NP, B(1)=Desc, E(4)=Costo, I(8)=Exist, J(9)=Ingreso, K(10)=Ult Comp, L(11)=Ult Vta
                            df_c['NP'] = df_r.iloc[:,0].astype(str).str.strip()
                            df_c['DESCRIPCION'] = df_r.iloc[:,1].astype(str).str.strip()
                            
                            # Numéricos
                            df_c['COSTO_UNITARIO'] = pd.to_numeric(df_r.iloc[:,4], errors='coerce').fillna(0)
                            df_c['EXISTENCIA'] = pd.to_numeric(df_r.iloc[:,8], errors='coerce').fillna(0)
                            
                            # Fechas (Solo guardamos la fecha, Power BI hará la matemática)
                            df_c['FEC_INGRESO'] = pd.to_datetime(df_r.iloc[:,9], errors='coerce').dt.date
                            df_c['FEC_ULT_COMPRA'] = pd.to_datetime(df_r.iloc[:,10], errors='coerce').dt.date
                            df_c['FEC_ULT_VENTA_INV'] = pd.to_datetime(df_r.iloc[:,11], errors='coerce').dt.date # Fecha reportada por sistema inventario
                            
                            # Etiquetas fijas
                            df_c['SUCURSAL'] = sucursal
                            df_c['ALMACEN'] = almacen
                            
                            inventarios_cargados.append(df_c)
                            st.success(f"✅ {len(df_c)} regs.")
                        except Exception as e:
                            st.error(f"❌ Error en {almacen}: Verifica estructura de columnas.")

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
                        if 'CATEGORIA' in df_final.columns:
                            df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna("X DEFINIR")
                        if 'LINEA' in df_final.columns:
                            df_final['LINEA'] = df_final['LINEA'].fillna("X DEFINIR")
                
                # 3. Limpieza final de Nulos
                df_final['EXISTENCIA'] = df_final['EXISTENCIA'].fillna(0)
                df_final['COSTO_UNITARIO'] = df_final['COSTO_UNITARIO'].fillna(0)
                df_final['VALOR_INVENTARIO'] = df_final['EXISTENCIA'] * df_final['COSTO_UNITARIO']
                
                # 4. Generar Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='DATA_INVENTARIO')
                    
                fecha_hoy = datetime.now().strftime('%Y-%m-%d')
                nombre_archivo = f"INVENTARIO_CONSOLIDADO_{fecha_hoy}.xlsx"
                
                st.success(f"¡Listo! Archivo generado con {len(df_final)} registros.")
                st.warning("⚠️ Guarda este archivo en: BI_REFACCIONES_MASTER > DATA RESOURCES > DATA_INVENTARIO")
                
                st.download_button(
                    label="📥 Descargar Excel Limpio",
                    data=buffer.getvalue(),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("No has cargado ningún inventario.")
            

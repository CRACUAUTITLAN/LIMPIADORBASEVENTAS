import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
from dateutil.relativedelta import relativedelta

def show():
    st.header("🏥 Generador Base Almacen (Salud de Inventario)")
    st.markdown("---")

    # ======================================================
    # 1. FUNCIONES
    # ======================================================
    def cargar_catalogo_blindado(uploaded_file):
        if uploaded_file is None: return None
        df = pd.read_excel(uploaded_file)
        df.rename(columns={df.columns[0]: 'NP'}, inplace=True)
        if 'NP' not in df.columns: df.rename(columns={'Numero de Parte': 'NP'}, inplace=True)
        df['NP'] = df['NP'].astype(str).str.strip()
        df = df.drop_duplicates(subset=['NP'], keep='first')
        return df[['NP', 'LINEA', 'CATEGORIA']]

    def calcular_salud_inventario(row, fecha_corte):
        meses_venta_10 = row['MESES_CON_VENTA_10M']
        meses_venta_12 = row['MESES_CON_VENTA_12M']
        
        if meses_venta_10 > 5: return "ALTO MOVIMIENTO"
        elif 3 <= meses_venta_10 <= 5: return "MEDIO MOVIMIENTO"
        elif 1 <= meses_venta_10 <= 2: return "BAJO MOVIMIENTO"
        elif meses_venta_10 == 0 and meses_venta_12 > 0: return "NULO / RIESGO"
        else: return "OBSOLETO"

    # ======================================================
    # 2. INTERFAZ: CONFIGURACIÓN
    # ======================================================
    col_conf1, col_conf2 = st.columns(2)
    with col_conf1:
        st.subheader("1. Configuración")
        file_cat = st.file_uploader("📂 Catálogo de Partes (.xlsx)", type=["xlsx"], key="inv_cat")
        file_map = st.file_uploader("🗺️ Mapa Sucursales (.xlsx)", type=["xlsx"], key="inv_map")
    
    with col_conf2:
        st.subheader("2. Historial Ventas")
        files_ventas = st.file_uploader("📊 Master Ventas (Varios)", accept_multiple_files=True, type=["xlsx"], key="inv_ventas")
        fecha_corte = st.date_input("📅 Fecha Corte", datetime.now())
        fecha_corte = pd.to_datetime(fecha_corte)

    df_frecuencia = pd.DataFrame()

    if files_ventas:
        with st.status("Procesando historial de ventas...") as status:
            all_ventas = []
            for f in files_ventas:
                try:
                    temp = pd.read_excel(f)
                    temp['FECHA_DT'] = pd.to_datetime(temp['FECHA'], dayfirst=True, errors='coerce')
                    temp['NP'] = temp['NP'].astype(str).str.strip()
                    all_ventas.append(temp[['NP', 'FECHA_DT']])
                except: pass
            
            if all_ventas:
                df_full = pd.concat(all_ventas)
                fecha_lim_10 = fecha_corte - relativedelta(months=10)
                fecha_lim_12 = fecha_corte - relativedelta(months=12)
                
                df_full = df_full[df_full['FECHA_DT'] >= fecha_lim_12]
                df_full['MES_ANIO'] = df_full['FECHA_DT'].dt.to_period('M')
                
                c_12 = df_full.groupby('NP')['MES_ANIO'].nunique().reset_index().rename(columns={'MES_ANIO': 'MESES_CON_VENTA_12M'})
                
                v_10 = df_full[df_full['FECHA_DT'] >= fecha_lim_10]
                c_10 = v_10.groupby('NP')['MES_ANIO'].nunique().reset_index().rename(columns={'MES_ANIO': 'MESES_CON_VENTA_10M'})
                
                df_frecuencia = pd.merge(c_12, c_10, on='NP', how='left').fillna(0)
                status.update(label="✅ Historial de ventas listo!", state="complete")

    # ======================================================
    # 3. CARGA DE INVENTARIOS
    # ======================================================
    st.markdown("---")
    st.subheader("3. Carga por Almacén")

    if file_map and file_cat and not df_frecuencia.empty:
        try:
            df_mapa = pd.read_excel(file_map)
            # Ajuste simple de columnas si no trae encabezado
            if "SUCURSAL" not in str(df_mapa.columns[0]).upper():
                df_mapa = pd.read_excel(file_map, header=None)
                df_mapa.columns = ['SUCURSAL', 'ALMACEN']
            else:
                df_mapa.columns = ['SUCURSAL', 'ALMACEN']

            sucursales = df_mapa['SUCURSAL'].unique()
            inventarios_cargados = []
            
            cols = st.columns(len(sucursales))
            for idx, suc in enumerate(sucursales):
                with cols[idx]:
                    st.markdown(f"#### 🏢 {suc}")
                    almacenes = df_mapa[df_mapa['SUCURSAL'] == suc]['ALMACEN'].unique()
                    for alm in almacenes:
                        f_inv = st.file_uploader(f"{alm}", type=["xls", "xlsx"], key=f"inv_{suc}_{alm}")
                        if f_inv:
                            df_r = pd.read_excel(f_inv, header=None)
                            df_c = pd.DataFrame()
                            # Mapeo Columnas A, B, E, I, J, K, L
                            df_c['NP'] = df_r.iloc[:,0].astype(str).str.strip()
                            df_c['DESCRIPCION'] = df_r.iloc[:,1].astype(str).str.strip()
                            df_c['COSTO'] = pd.to_numeric(df_r.iloc[:,4], errors='coerce').fillna(0)
                            df_c['EXISTENCIA'] = pd.to_numeric(df_r.iloc[:,8], errors='coerce').fillna(0)
                            df_c['INGRESO'] = pd.to_datetime(df_r.iloc[:,9], errors='coerce').dt.date
                            df_c['ULT_COMP'] = pd.to_datetime(df_r.iloc[:,10], errors='coerce').dt.date
                            df_c['ULT_VTA'] = pd.to_datetime(df_r.iloc[:,11], errors='coerce').dt.date
                            df_c['SUCURSAL'] = suc
                            df_c['ALMACEN'] = alm
                            inventarios_cargados.append(df_c)
                            st.success(f"Cargado: {len(df_c)} items")

            # ======================================================
            # 4. GENERAR
            # ======================================================
            if st.button("🚀 Procesar Inventario Maestro", type="primary", use_container_width=True):
                if inventarios_cargados:
                    with st.spinner("Calculando obsolescencia..."):
                        df_fin = pd.concat(inventarios_cargados, ignore_index=True)
                        df_cat = cargar_catalogo_blindado(file_cat)
                        
                        df_fin = pd.merge(df_fin, df_cat, on='NP', how='left').fillna("X DEFINIR")
                        df_fin = pd.merge(df_fin, df_frecuencia, on='NP', how='left').fillna(0)
                        
                        df_fin['ESTATUS'] = df_fin.apply(lambda x: calcular_salud_inventario(x, fecha_corte), axis=1)
                        
                        # Exportar
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df_fin.to_excel(writer, index=False)
                            # Formato colores
                            wb = writer.book
                            ws = writer.sheets['Sheet1']
                            f_red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                            f_grn = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
                            
                            col_stat = df_fin.columns.get_loc('ESTATUS')
                            ws.conditional_format(1, col_stat, len(df_fin), col_stat, {'type': 'text', 'criteria': 'containing', 'value': 'OBSOLETO', 'format': f_red})
                            ws.conditional_format(1, col_stat, len(df_fin), col_stat, {'type': 'text', 'criteria': 'containing', 'value': 'ALTO', 'format': f_grn})
                        
                        st.download_button("📥 Descargar Excel", buffer.getvalue(), f"Salud_Inv_{fecha_corte.date()}.xlsx")
                else:
                    st.error("Sube al menos un inventario.")
                    
        except Exception as e:
            st.error(f"Error en archivos de configuración: {e}")
    else:
        st.info("👆 Carga el catálogo, mapa y ventas para habilitar los almacenes.")

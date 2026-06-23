import streamlit as st
import pandas as pd
import io

def procesar_auditoria(df_compras, df_ventas):
    # 1. Estandarizar nombres de columnas (quitar espacios extra y mayúsculas)
    df_compras.columns = df_compras.columns.str.strip().str.upper()
    df_ventas.columns = df_ventas.columns.str.strip().str.upper()

    # 2. Agrupar Base de Compras (Sumamos cantidad, conservamos costo y fecha max)
    compras_grp = df_compras.groupby('NP').agg(
        DESCRIPCION=('DESCRIPCION', 'first'),
        FECHA_ULT_COMPRA=('FECHA', 'max'),
        CANTIDAD_COMPRADA=('CANTIDAD', 'sum'),
        COSTO_UNITARIO=('COSTO UNITARIO', 'max') # Tomamos el costo más alto o reciente
    ).reset_index()

    # 3. Agrupar Base de Ventas (Sumamos cantidad vendida en el mes en todas las sucursales)
    ventas_grp = df_ventas.groupby('NP').agg(
        CANTIDAD_VENDIDA=('CANTIDAD', 'sum')
    ).reset_index()

    # 4. Cruce Maestro (Left Join desde compras, porque queremos auditar lo que se compró)
    df_final = pd.merge(compras_grp, ventas_grp, on='NP', how='left')
    df_final['CANTIDAD_VENDIDA'] = df_final['CANTIDAD_VENDIDA'].fillna(0)

    # 5. Cálculos de Eficiencia y Congelamiento
    df_final['% DE VENTA'] = (df_final['CANTIDAD_VENDIDA'] / df_final['CANTIDAD_COMPRADA'])
    
    # Si se vendió más de lo comprado (por stock pasado), la cantidad congelada es 0
    df_final['CANTIDAD CONGELADA'] = df_final.apply(lambda x: max(0, x['CANTIDAD_COMPRADA'] - x['CANTIDAD_VENDIDA']), axis=1)
    
    df_final['TOTAL INVERTIDO ($)'] = df_final['CANTIDAD_COMPRADA'] * df_final['COSTO_UNITARIO']
    df_final['CAPITAL CONGELADO ($)'] = df_final['CANTIDAD CONGELADA'] * df_final['COSTO_UNITARIO']

    # 6. Ordenar por el mayor Capital Congelado (De peor a mejor)
    df_final = df_final.sort_values('CAPITAL CONGELADO ($)', ascending=False)
    
    # Reordenar columnas para una vista limpia
    df_final = df_final[['NP', 'DESCRIPCION', 'FECHA_ULT_COMPRA', 'COSTO_UNITARIO', 'CANTIDAD_COMPRADA', 'CANTIDAD_VENDIDA', '% DE VENTA', 'CANTIDAD CONGELADA', 'TOTAL INVERTIDO ($)', 'CAPITAL CONGELADO ($)']]

    return df_final

def generar_excel_kpi(df_final):
    output = io.BytesIO()
    
    # Usamos openpyxl para crear múltiples hojas
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- HOJA 1: DASHBOARD KPI ---
        total_invertido = df_final['TOTAL INVERTIDO ($)'].sum()
        total_congelado = df_final['CAPITAL CONGELADO ($)'].sum()
        eficiencia = 1 - (total_congelado / total_invertido) if total_invertido > 0 else 0
        
        kpi_data = {
            'KPI ESTRATÉGICO': ['💰 Total Invertido en Compras del Mes', '🧊 Capital Congelado (Sin Venta)', '⚡ Eficiencia de Compra del Mes (%)'],
            'RESULTADO': [total_invertido, total_congelado, eficiencia]
        }
        df_kpi = pd.DataFrame(kpi_data)
        df_kpi.to_excel(writer, sheet_name='Dashboard KPI', index=False, startrow=1, startcol=1)
        
        # Agregamos el TOP 10 de peores compras debajo de los KPIs
        top_10 = df_final.head(10)[['NP', 'DESCRIPCION', 'CANTIDAD CONGELADA', 'CAPITAL CONGELADO ($)']]
        top_10.to_excel(writer, sheet_name='Dashboard KPI', index=False, startrow=6, startcol=1)
        
        # --- HOJA 2: DETALLE COMPLETO ---
        df_final.to_excel(writer, sheet_name='Detalle Analítico', index=False)

    return output.getvalue()

def show():
    st.title("💸 Análisis de Eficiencia de Compras y Capital Congelado")
    st.markdown("Sube las bases de datos de **Ventas** y **Compras** filtradas al mes en curso para generar la auditoría de compras.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Base de Compras")
        file_compras = st.file_uploader("Sube el Excel de Compras (Agencia, Factura, Fecha...)", type=['xlsx', 'xls'], key="compra")
    with col2:
        st.subheader("2. Base de Ventas")
        file_ventas = st.file_uploader("Sube el Excel de Ventas (Sucursal, Cliente, NP...)", type=['xlsx', 'xls'], key="venta")

    if file_compras and file_ventas:
        if st.button("🚀 Ejecutar Análisis", type="primary", use_container_width=True):
            try:
                with st.spinner('Analizando millones de pesos...'):
                    df_compras = pd.read_excel(file_compras)
                    df_ventas = pd.read_excel(file_ventas)
                    
                    df_resultado = procesar_auditoria(df_compras, df_ventas)
                    excel_data = generar_excel_kpi(df_resultado)
                
                st.success("✅ ¡Análisis completado con éxito!")
                
                # Mostrar KPIs rápidos en la pantalla
                st.markdown("### 📊 Resumen Rápido")
                m1, m2, m3 = st.columns(3)
                total_inv = df_resultado['TOTAL INVERTIDO ($)'].sum()
                total_cong = df_resultado['CAPITAL CONGELADO ($)'].sum()
                efi = 1 - (total_cong / total_inv) if total_inv > 0 else 0
                
                m1.metric("Total Invertido", f"${total_inv:,.2f}")
                m2.metric("Capital Congelado", f"${total_cong:,.2f}")
                m3.metric("Eficiencia de Compra", f"{efi*100:.1f}%")

                st.download_button(
                    label="📥 Descargar Excel Final (Dashboard + Detalle)",
                    data=excel_data,
                    file_name="Auditoria_Compras_KPI.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al procesar los archivos: {e}")

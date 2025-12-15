import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# Configuración de la página
st.set_page_config(page_title="Limpieza Inventario Pro", layout="wide")

st.title("🛠️ Limpieza de Base de Datos (Lógica por Fechas)")
st.markdown("""
**Reglas Actualizadas:**
1. **Factura:** Se identifica si la Columna C tiene una FECHA válida.
2. **Cliente:** Se ubica arriba de la primera factura del bloque y termina en "Totales de:".
3. **Orden:** Cronológico (Fecha) -> Alfabético (Cliente).
""")

# --- FUNCIONES DE AYUDA ---

def is_valid_date(value):
    """
    Verifica si un valor es una fecha válida (objeto datetime o string dd/mm/yyyy).
    Retorna True si es fecha, False si no.
    """
    if pd.isna(value):
        return False
    
    # Caso 1: Ya es un objeto datetime (Excel lo leyó como fecha)
    if isinstance(value, (datetime, pd.Timestamp)):
        return True
    
    # Caso 2: Es un string, intentamos parsear formato dd/mm/yyyy
    if isinstance(value, str):
        try:
            # Intentar formatos comunes
            pd.to_datetime(value, dayfirst=True)
            # Verificación extra: que parezca fecha (contiene / o -)
            if '/' in value or '-' in value:
                return True
        except:
            return False
            
    return False

def process_excel_file(uploaded_file):
    """
    Lee todas las hojas y aplica la lógica basada en la Columna C (Fecha)
    """
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    
    all_clean_rows = []

    # Iterar sobre cada hoja
    for sheet_name, df in xls.items():
        data = df.values.tolist()
        
        # Variables de estado (Memoria)
        current_client = None
        current_fecha = None
        current_tipo = None
        current_almacen = None
        
        for i, row in enumerate(data):
            # Saltamos filas vacías
            if len(row) < 3:
                continue
                
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = row[1]
            col_c = row[2] # Aquí buscamos la fecha
            
            # --- 1. DETECTOR DE FIN DE CLIENTE ---
            if col_a.startswith("Totales de:"):
                current_client = None
                continue
            
            # --- 2. DETECTOR DE FACTURA (REGLA DE ORO: COL C ES FECHA) ---
            if is_valid_date(col_c):
                # Es una línea de factura
                # Actualizamos datos de memoria
                current_fecha = col_c
                
                # Ubicación de datos según instrucción previa:
                # FECHA: Col C (indice 2)
                # TIPO: Col D (indice 3)
                # ALMACEN: Col E (indice 4)
                if len(row) > 4:
                    current_tipo = row[3]
                    current_almacen = row[4]
                
                # --- BUSQUEDA DE CLIENTE ---
                # Si acabamos de encontrar una factura y NO tenemos cliente (es el inicio del bloque)
                # buscamos el nombre 2 filas arriba (estándar visual)
                if current_client is None:
                    # Verificar límites para no dar error
                    if i >= 2:
                        posible_cliente = str(data[i-2][0]).strip()
                        # Limpiamos prefijos comunes si existen
                        current_client = posible_cliente.replace("- - ", "").strip()
                    elif i >= 1: # Intento de respaldo 1 fila arriba
                        current_client = str(data[i-1][0]).replace("- - ", "").strip()
                
                continue # Terminamos con esta fila, pasamos a la siguiente

            # --- 3. DETECTOR DE PRODUCTO (ID PART) ---
            # Condiciones:
            # - No es factura (ya pasó el check arriba)
            # - Columna A tiene texto (el ID)
            # - Columna A NO es "ID Parte" (encabezado basura)
            # - Tenemos un cliente y una fecha activos
            
            if col_a != "" and col_a != "ID Parte" and col_a != "nan":
                # Validar que Columna B sea texto (Descripción) para estar seguros
                # Si B es número, podría ser basura, pero confiaremos en que si A tiene ID y no es factura, es prod.
                
                # Asegurarnos de tener contexto (Cliente y Fecha detectados previamente)
                if current_client and current_fecha:
                     if len(row) > 10:
                        clean_row = {
                            "CLIENTE": current_client,
                            "PART ID": col_a,
                            "DESCR": str(col_b), # Forzar string
                            "LINEA": row[2],     # En producto, Linea está en C
                            "FECHA": current_fecha, # Heredado de la factura
                            "TIPO_FACTURA": current_tipo, # Heredado
                            "ALMACEN": current_almacen,   # Heredado
                            "CANTIDAD": row[3],  # En producto, Cantidad está en D
                            "COSTO DE VENTA": row[6],
                            "VENTA SUBTOTAL": row[7],
                            "IVA": row[8],
                            "VENTA TOTAL": row[9],
                            "UTILIDAD": row[10]
                        }
                        all_clean_rows.append(clean_row)

    # Crear DataFrame
    df_result = pd.DataFrame(all_clean_rows)
    return df_result

# --- INTERFAZ ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Cargar Reporte Sucio (.xls)")
    file_sucio = st.file_uploader("Sube el archivo XLS", type=["xls"])

with col2:
    st.subheader("2. Cargar Actualizador (.xlsx)")
    file_actualizador = st.file_uploader("Sube el archivo de Categorías", type=["xlsx", "xls"])

if st.button("Ejecutar Limpieza y Ordenamiento"):
    if file_sucio is not None and file_actualizador is not None:
        try:
            with st.spinner('Procesando múltiples hojas y aplicando reglas de fecha...'):
                # 1. Limpieza
                df_limpio = process_excel_file(file_sucio)
                
                if df_limpio.empty:
                    st.error("No se encontraron registros. Verifica que la columna C tenga fechas en las filas de factura.")
                else:
                    st.success(f"Registros extraídos: {len(df_limpio)}")
                    
                    # 2. Cargar Actualizador
                    df_update = pd.read_excel(file_actualizador)
                    
                    # Normalizar columnas de cruce
                    df_limpio['PART ID'] = df_limpio['PART ID'].astype(str).str.strip()
                    df_update['NP'] = df_update['NP'].astype(str).str.strip()
                    
                    # 3. Cruce (Merge)
                    df_final = pd.merge(
                        df_limpio, 
                        df_update[['NP', 'LINEA', 'CATEGORIA']], 
                        left_on='PART ID', 
                        right_on='NP', 
                        how='left',
                        suffixes=('_orig', '_new')
                    )
                    
                    # 4. Actualización condicional
                    df_final['LINEA'] = np.where(df_final['LINEA_new'].notna(), df_final['LINEA_new'], df_final['LINEA_orig'])
                    df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                    
                    # 5. ORDENAMIENTO (NUEVO REQUERIMIENTO)
                    # Convertir FECHA a datetime real para ordenar cronológicamente
                    df_final['FECHA'] = pd.to_datetime(df_final['FECHA'], dayfirst=True, errors='coerce')
                    
                    # Ordenar: 1. Fecha (Asc), 2. Cliente (Asc)
                    df_final = df_final.sort_values(by=['FECHA', 'CLIENTE'], ascending=[True, True])
                    
                    # Formatear la fecha a texto limpio (dd/mm/yyyy) para el excel final
                    df_final['FECHA'] = df_final['FECHA'].dt.strftime('%d/%m/%Y')

                    # 6. Selección de columnas finales
                    cols_order = [
                        "CLIENTE", "PART ID", "DESCR", "LINEA", "CATEGORIA", 
                        "FECHA", "TIPO_FACTURA", "ALMACEN", "CANTIDAD", 
                        "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"
                    ]
                    df_final = df_final[cols_order]
                    
                    st.dataframe(df_final.head())
                    
                    # 7. Descarga
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Base Maestra 2024-2025')
                    
                    st.download_button(
                        label="📥 Descargar Base Final Ordenada",
                        data=buffer.getvalue(),
                        file_name="Reporte_Final_Ordenado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error técnico: {e}")
    else:
        st.info("Sube ambos archivos para comenzar.")

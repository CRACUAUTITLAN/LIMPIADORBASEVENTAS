import streamlit as st
import pandas as pd
import numpy as np
import io

# Configuración de la página
st.set_page_config(page_title="Limpieza y Actualización de Inventario", layout="wide")

st.title("🛠️ Limpieza de Base de Datos y Actualizador de Líneas")
st.markdown("""
Esta aplicación procesa reportes masivos en formato **.xls** (múltiples hojas), limpia la estructura 
y actualiza la información de **Línea** y **Categoría** cruzando con una base maestra.
""")

# --- FUNCIONES DE AYUDA ---

def is_number(s):
    """Devuelve True si el valor puede ser un número, False si es texto."""
    try:
        float(str(s))
        return True
    except ValueError:
        return False

def process_excel_file(uploaded_file):
    """
    Lee todas las hojas del archivo .xls sucio y aplica la lógica de extracción.
    """
    # Leer todas las hojas (sheet_name=None devuelve un diccionario de dataframes)
    # Usamos header=None para manejar nosotros las posiciones
    xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    
    all_clean_rows = []

    # Iterar sobre cada hoja del archivo
    for sheet_name, df in xls.items():
        # Convertir a lista de listas para iteración rápida
        data = df.values.tolist()
        
        # Variables de estado (Memoria)
        current_client = None
        current_fecha = None
        current_tipo = None
        current_almacen = None
        
        # Lógica de arranque: Encontrar el primer cliente
        # Buscamos el primer "ID Parte" y subimos 2 filas
        start_row_index = 0
        for i, row in enumerate(data):
            if len(row) > 0 and str(row[0]).strip() == "ID Parte":
                # Validar que no nos salgamos del índice hacia arriba
                if i >= 2:
                    current_client = str(data[i-2][0]).replace("- - ", "").strip()
                start_row_index = i # Empezamos a procesar desde aquí
                break
        
        # Iterar fila por fila
        for i in range(start_row_index, len(data)):
            row = data[i]
            
            # Verificar si la fila está vacía o muy corta
            if len(row) < 2:
                continue
                
            col_a = str(row[0]).strip() if pd.notna(row[0]) else ""
            col_b = row[1] # Mantenemos el tipo original para checar si es numero
            
            # --- LÓGICA DE CLIENTES ---
            # Si encontramos "Totales de:", el siguiente cliente está en la fila de abajo
            if col_a.startswith("Totales de:"):
                if i + 1 < len(data):
                    raw_client = str(data[i+1][0])
                    # Limpieza básica del nombre (quitar guiones si existen)
                    current_client = raw_client.replace("- - ", "").strip()
                continue
            
            # --- IGNORAR ENCABEZADOS REPETITIVOS ---
            if col_a == "ID Parte" or col_a == "nan" or col_a == "":
                continue

            # --- LÓGICA DE FACTURAS VS PRODUCTOS ---
            # Revisamos la Columna B para decidir
            
            b_is_num = is_number(col_b)
            
            if b_is_num:
                # ==> ES FILA DE FACTURA
                # Actualizamos variables de estado (Fecha, Tipo, Almacen)
                # Según instrucciones: 
                # Fecha: Col C (indice 2)
                # Tipo: Col D (indice 3)
                # Almacen: Col E (indice 4)
                if len(row) > 4:
                    current_fecha = row[2]
                    current_tipo = row[3]
                    current_almacen = row[4]
            else:
                # ==> ES FILA DE PRODUCTO (ID PART)
                # Col A: ID PART
                # Col B: DESCR (Texto)
                # Col C: LINEA
                # Col D: CANTIDAD
                # Col G: COSTO VENTA (indice 6)
                # Col H: VENTA SUBTOTAL (indice 7)
                # Col I: IVA (indice 8)
                # Col J: VENTA TOTAL (indice 9)
                # Col K: UTILIDAD (indice 10)
                
                # Validamos que tengamos suficientes columnas para evitar errores
                if len(row) > 10:
                    clean_row = {
                        "CLIENTE": current_client,
                        "PART ID": col_a,
                        "DESCR": col_b,
                        "LINEA": row[2],
                        "FECHA": current_fecha,
                        "TIPO_FACTURA": current_tipo, # Renombrado para evitar conflicto con tipo de cambio/producto
                        "ALMACEN": current_almacen,
                        "CANTIDAD": row[3], # OJO: En fila de producto, Cantidad está en Col D
                        "COSTO DE VENTA": row[6],
                        "VENTA SUBTOTAL": row[7],
                        "IVA": row[8],
                        "VENTA TOTAL": row[9],
                        "UTILIDAD": row[10]
                    }
                    all_clean_rows.append(clean_row)

    # Crear DataFrame final
    df_result = pd.DataFrame(all_clean_rows)
    return df_result

# --- INTERFAZ DE USUARIO ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Cargar Reporte Sucio (.xls)")
    file_sucio = st.file_uploader("Sube el archivo XLS (Multi-hojas)", type=["xls"])

with col2:
    st.subheader("2. Cargar Actualizador (.xlsx)")
    file_actualizador = st.file_uploader("Sube el archivo con NP, LINEA, CATEGORIA", type=["xlsx", "xls"])

if st.button("Procesar Archivos"):
    if file_sucio is not None and file_actualizador is not None:
        try:
            with st.spinner('Limpiando base de datos masiva (esto puede tardar un poco)...'):
                # 1. Procesar archivo sucio
                df_limpio = process_excel_file(file_sucio)
                st.success(f"Base limpiada con éxito: {len(df_limpio)} registros encontrados.")
            
            with st.spinner('Cruzando con el actualizador...'):
                # 2. Cargar actualizador
                df_update = pd.read_excel(file_actualizador)
                
                # Asegurar que las columnas existan y estén limpias
                # Convertimos a string para asegurar match perfecto
                df_limpio['PART ID'] = df_limpio['PART ID'].astype(str).str.strip()
                df_update['NP'] = df_update['NP'].astype(str).str.strip()
                
                # 3. Merge (Cruce)
                # Hacemos un Left Join: Mantenemos todo lo del reporte limpio y traemos datos del update
                df_final = pd.merge(
                    df_limpio, 
                    df_update[['NP', 'LINEA', 'CATEGORIA']], 
                    left_on='PART ID', 
                    right_on='NP', 
                    how='left',
                    suffixes=('_orig', '_new')
                )
                
                # 4. Lógica de Reemplazo (np.where)
                # Si 'LINEA_new' (del actualizador) existe, úsala. Si no, usa 'LINEA_orig'.
                df_final['LINEA'] = np.where(
                    df_final['LINEA_new'].notna(), 
                    df_final['LINEA_new'], 
                    df_final['LINEA_orig']
                )
                
                # Llenar CATEGORIA con 'Desconocido' si no hubo match
                df_final['CATEGORIA'] = df_final['CATEGORIA'].fillna('Desconocido')
                
                # 5. Reordenar y limpiar columnas
                # Definimos el orden solicitado
                cols_order = [
                    "CLIENTE", "PART ID", "DESCR", "LINEA", "CATEGORIA", 
                    "FECHA", "TIPO_FACTURA", "ALMACEN", "CANTIDAD", 
                    "COSTO DE VENTA", "VENTA SUBTOTAL", "IVA", "VENTA TOTAL", "UTILIDAD"
                ]
                
                # Seleccionamos solo las columnas finales
                df_final = df_final[cols_order]
                
                st.success("¡Cruce y actualización completados!")
                st.dataframe(df_final.head())
                
                # 6. Botón de descarga
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Base Limpia')
                
                st.download_button(
                    label="📥 Descargar Excel Final",
                    data=buffer.getvalue(),
                    file_name="Reporte_Limpio_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
    else:
        st.warning("Por favor sube ambos archivos para continuar.")

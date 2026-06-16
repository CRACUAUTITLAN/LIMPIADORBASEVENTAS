import streamlit as st
import ventas      # Importamos tu archivo de ventas
import inventario  # Importamos tu archivo de inventario
import fill_rate   # Importamos el nuevo limpiador de Fill Rate
import almacen_actual_idealease # 🔴 NUEVO: Importamos el limpiador de Almacén Actual

# CONFIGURACIÓN GENERAL
st.set_page_config(page_title="Modelo BI Master", layout="wide")

# ==========================
# 🔐 SISTEMA DE LOGIN
# ==========================
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False
if 'usuario_activo' not in st.session_state:
    st.session_state['usuario_activo'] = ""

CREDENTIALS = {
    "Eduardo Barrera": "Barrera1234",
    "Ivonne Ascencio": "Ascencio123.",
    "Luis Adrian" : "Luis1234."
}

def check_login(user, password):
    if CREDENTIALS.get(user) == password:
        st.session_state['authenticated'] = True
        st.session_state['usuario_activo'] = user
        st.rerun()
    else:
        st.error("🚫 Contraseña incorrecta")

if not st.session_state['authenticated']:
    st.title("🔒 Acceso Corporativo")
    c1, c2 = st.columns([1,2])
    with c1:
        user = st.selectbox("Usuario", list(CREDENTIALS.keys()))
        pw = st.text_input("Contraseña", type="password")
        if st.button("Entrar"):
            check_login(user, pw)
    st.stop()

# ==========================
# ☰ MENÚ LATERAL (SIDEBAR)
# ==========================
with st.sidebar:
    st.title("🗂️ Menú Principal")
    usuario = st.session_state['usuario_activo']
    st.write(f"Hola, **{usuario}**")
    
    if usuario == "Ivonne Ascencio":
        opciones_disponibles = ["Generador Base Almacen"]
    else:
        # 🔴 NUEVO: Agregamos "ALMACEN ACTUAL IDEALEASE" a la lista
        opciones_disponibles = [
            "Generador Base Ventas", 
            "Generador Base Almacen", 
            "Generador Base Fill Rate",
            "ALMACEN ACTUAL IDEALEASE"
        ]
    
    opcion = st.radio(
        "Selecciona Herramienta:",
        opciones_disponibles
    )
    
    st.markdown("---")
    if st.button("Cerrar Sesión"):
        st.session_state['authenticated'] = False
        st.rerun()

# ==========================
# 🚀 CONTROLADOR DE VISTAS
# ==========================
if opcion == "Generador Base Ventas":
    ventas.show()  
elif opcion == "Generador Base Almacen":
    inventario.show()
elif opcion == "Generador Base Fill Rate":
    fill_rate.show() 
elif opcion == "ALMACEN ACTUAL IDEALEASE":
    almacen_actual_idealease.show() # 🔴 NUEVO: Llamamos a la vista

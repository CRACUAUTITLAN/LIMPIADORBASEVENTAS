import streamlit as st
import ventas      # Importamos tu archivo de ventas
import inventario  # Importamos tu archivo de inventario

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
    "Manuel Salas": "Salas1234",
    "Eduardo Barrera": "Barrera1234"
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
    st.write(f"Hola, **{st.session_state['usuario_activo']}**")
    
    opcion = st.radio(
        "Selecciona Herramienta:",
        ["Generador Base Ventas", "Generador Base Almacen"]
    )
    
    st.markdown("---")
    if st.button("Cerrar Sesión"):
        st.session_state['authenticated'] = False
        st.rerun()

# ==========================
# 🚀 CONTROLADOR DE VISTAS
# ==========================
if opcion == "Generador Base Ventas":
    ventas.show()  # Llamamos a la función show() dentro de ventas.py

elif opcion == "Generador Base Almacen":
    inventario.show() # Llamamos a la función show() dentro de inventario.py

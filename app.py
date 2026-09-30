import streamlit as st

# 1. Configuración de la página
st.set_page_config(
    page_title="Mi Video",
    page_icon="🎬",
    layout="centered" # Mantiene el video centrado y con un buen tamaño
)

# 2. Título o encabezado (opcional, puedes cambiar el texto)
st.title("")
st.write("")

# 3. Reproductor de video preestablecido
# AQUÍ DEBES PONER EL NOMBRE EXACTO DE TU ARCHIVO DE VIDEO
st.video("mivideo.mp4")

import streamlit as st
APP_NAME = st.secrets.get("app", {}).get("name", "COMEX PDF READER")
PAGES = ["Home", "Aplicación Comex", "Aplicación Archivo Gastos", "Configurações"]

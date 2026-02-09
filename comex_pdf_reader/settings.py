import streamlit as st
APP_NAME = st.secrets.get("app", {}).get("name", "COMEX PDF READER")
PAGES = ["Home", "Processar PDFs", "Configurações"]

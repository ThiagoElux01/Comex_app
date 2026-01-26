
import streamlit as st

def render():
    st.subheader("Bem-vindo(a)! 👋")
    st.info("Projeto piloto de extração de informações de pdfs, COMEX PDF READER")
    st.markdown(
        """
        - Upload de PDFs
        - Extração de texto/tabelas
        - Exportação para Excel
        - Integração com Tasa SUNAT
        """
    )

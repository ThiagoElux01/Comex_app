

import streamlit as st
import pandas as pd
from pathlib import Path

# Caminho do arquivo modelo
ARQUIVO_MODELO = Path("assets/modelos/Externos.xlsx")

@st.cache_data
def carregar_modelo():
    return pd.read_excel(ARQUIVO_MODELO)

def render():
    st.subheader("Home")

    st.write("Atualização do arquivo de modelos externos")

    if st.button("🔄 Update"):
        try:
            st.cache_data.clear()
            df = carregar_modelo()

            st.success("Arquivo atualizado com sucesso ✅")
            st.info(f"📊 Total de linhas: {len(df)}")

        except Exception as e:
            st.error(f"Erro ao atualizar arquivo: {e}")

    # opcional: preview
    with st.expander("Ver prévia dos dados"):
        try:
            df = carregar_modelo()
            st.dataframe(df.head(10), use_container_width=True)
        except:
            st.warning("Arquivo ainda não carregado.")


import streamlit as st
import pandas as pd
from pathlib import Path

# raiz do projeto (…/comex_pdf_reader)
BASE_DIR = Path(__file__).resolve().parents[2]

# caminho correto do arquivo
ARQUIVO_MODELO = BASE_DIR / "assets" / "modelos" / "Externos.xlsx"

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

    # debug opcional (pode remover depois)
    with st.expander("ℹ️ Caminho do arquivo"):
        st.code(str(ARQUIVO_MODELO))

    with st.expander("Ver prévia dos dados"):
        try:
            df = carregar_modelo()
            st.dataframe(df.head(10), use_container_width=True)
        except:
            st.warning("Arquivo ainda não carregado.")

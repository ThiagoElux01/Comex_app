
# ui/pages/downloads_page.py
import streamlit as st
from pathlib import Path

ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets" / "modelos"

def _read_zip_bytes(filename: str) -> bytes:
    path = ASSETS_DIR / filename
    if not path.exists():
        st.error(f"Arquivo não encontrado: {path}")
        return b""
    return path.read_bytes()

def render():
    st.subheader("📦 Arquivos modelo")
    st.caption("Baixe os pacotes de templates para preparar os dados.")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            label="⬇️ Carga DUAS",
            data=_read_zip_bytes("carga_duas.zip"),
            file_name="carga_duas.zip",
            mime="application/zip",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="⬇️ Carga Externos",
            data=_read_zip_bytes("carga_externos.zip"),
            file_name="carga_externos.zip",
            mime="application/zip",
            use_container_width=True,
        )

    with col3:
        st.download_button(
            label="⬇️ Carga Adicionales",
            data=_read_zip_bytes("carga_adicionales.zip"),
            file_name="carga_adicionales.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.divider()
    st.info(
        "Dúvidas sobre o conteúdo dos modelos? Vá em **Processar PDFs** e veja os campos "
        "esperados em cada fluxo."
    )

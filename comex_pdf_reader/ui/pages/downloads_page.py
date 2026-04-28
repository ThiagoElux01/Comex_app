# ui/pages/downloads_page.py
import streamlit as st
from pathlib import Path

ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets" / "modelos"

def _read_file_bytes(filename: str) -> bytes:
    path = ASSETS_DIR / filename
    if not path.exists():
        st.error(f"Arquivo não encontrado: {path}")
        return b""
    return path.read_bytes()

def render():
    st.subheader("📦 Arquivos modelo")
    st.caption("Baixe os templates em Excel (.xlsx) para preparar os dados.")

    # Linha única com 4 colunas
    col4, col1, col2, col3 = st.columns(4)

    with col1:
        st.download_button(
            label="⬇️ Carga DUAS (XLSX)",
            data=_read_file_bytes("carga_duas.xlsx"),
            file_name="carga_duas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    with col2:
        st.download_button(
            label="⬇️ Carga Externos (XLSX)",
            data=_read_file_bytes("carga_externos.xlsx"),
            file_name="carga_externos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    # ⬇️ Empilhados na mesma coluna (fica um abaixo do outro)
    with col3:
        st.download_button(
            label="⬇️ Carga Adicionales 281110 (XLSX)",
            data=_read_file_bytes("carga_adicionales_10.xlsx"),
            file_name="carga_adicionales_10.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.write("")  # pequeno espaçador vertical
        st.download_button(
            label="⬇️ Carga Adicionales 281130 (XLSX)",
            data=_read_file_bytes("carga_adicionales_30.xlsx"),
            file_name="carga_adicionales_30.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    with col4:
        st.download_button(
            label="⬇️ Comex Report (XLSX)",
            data=_read_file_bytes("comex.xlsx"),
            file_name="comex.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    st.divider()
    st.info("Nesta página é possível baixar os arquivos modelo esperados por cada fluxo.")

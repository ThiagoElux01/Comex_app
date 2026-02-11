# ui/pages/downloads_page.py
import streamlit as st
from pathlib import Path
import datetime

ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets" / "modelos"

def _read_file_bytes(filename: str) -> bytes:
    path = ASSETS_DIR / filename
    if not path.exists():
        st.error(f"Arquivo não encontrado: {path}")
        return b""
    return path.read_bytes()

def render():
    # --- DIAGNÓSTICO (temporário) ---
    st.caption(f"downloads_page em: {Path(__file__).resolve()}")
    st.caption(f"ASSETS_DIR: {ASSETS_DIR}")
    try:
        st.caption("Arquivos em assets/modelos:")
        st.write(sorted([p.name for p in ASSETS_DIR.iterdir()]))
        st.caption(f"Build time: {datetime.datetime.now()}")
    except Exception as e:
        st.warning(f"Falha ao listar assets: {e}")
    # --- FIM DIAGNÓSTICO ---

    st.subheader("📦 Arquivos modelo")
    st.caption("Baixe os templates em Excel (.xlsx) para preparar os dados.")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            label="⬇️ Carga DUAS (XLSX) v2",
            data=_read_file_bytes("carga_duas.xlsx"),
            file_name="carga_duas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="⬇️ Carga Externos (XLSX) v2",
            data=_read_file_bytes("carga_externos.xlsx"),
            file_name="carga_externos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col3:
        st.download_button(
            label="⬇️ Carga Adicionales (XLSX) v2",
            data=_read_file_bytes("carga_adicionales.xlsx"),
            file_name="carga_adicionales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.info(
        "Nesta página é possível baixar os modelos em Excel esperados por cada fluxo."
    )

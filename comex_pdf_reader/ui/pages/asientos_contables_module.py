# -*- coding: utf-8 -*-
import re
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# Parser ASIENTOS CONTABLES – colunas fixas (GL - texto)
# -----------------------------------------------------------------------------
def parse_asientos_contables_fixed(texto: str):
    """Retorna registros columnados conforme layout de Thiago."""
    rows = []
    for ln in texto.splitlines():
        if not ln.strip():
            continue
        s = ln.rstrip('
')
        s_pad = s + ' ' * max(0, 92 - len(s))
        CC    = s_pad[0:2].strip()
        PROD  = s_pad[2:7].strip()
        CNT   = s_pad[7:14].strip()
        TDW   = s_pad[14:23].strip()
        Fcha  = s_pad[23:32].strip()
        Ntran = s_pad[32:41].strip()
        Debe  = s_pad[41:52].strip()
        Haber = s_pad[52:72].strip()
        Saldo = s_pad[72:92].strip()
        Texto = s[92:].strip() if len(s) > 92 else ''
        # heurística leve: aceita se tem data dd/mm/aa ou nro transação longo
        if not re.search(r"\d{2}/\d{2}/\d{2}", s):
            if not re.search(r"\d{6,}", Ntran):
                continue
        rows.append({
            'Cuenta': '',           # em branco
            'CC': CC,
            'PROD': PROD,
            'CNT': CNT,
            'TDW': TDW,
            'Fcha': Fcha,
            'Ntran': Ntran,
            'Debe': Debe,
            'Haber': Haber,
            'Saldo Real': '',       # em branco
            'Saldo': Saldo,
            'Texto': Texto,
        })
    return rows

# -----------------------------------------------------------------------------
# UI helper para integrar na "Aplicación Archivo Gastos"
# -----------------------------------------------------------------------------
def render_asientos_contables_ui(session_key_df: str = "aag_asientos_df"):
    st.subheader("📑 Asientos Contables (TXT)")
    st.caption(
        "Layout fijo: CC 1–2, PROD 3–7, CNT 8–14, TDW 15–23, Fcha 24–32, "
        "Ntran 33–41, Debe 42–52, Haber 53–72, Saldo 73–92, Texto 93–fin. "
        "Columnas nuevas 'Cuenta' y 'Saldo Real' quedan en blanco por ahora."
    )
    uploaded = st.file_uploader(
        "Seleccionar archivo (.txt)", type=["txt"], accept_multiple_files=False,
        key="asientos_upl",
    )
    col_run, col_clear = st.columns([2,1])
    with col_run:
        run_clicked = st.button("▶️ Ejecutar", type="primary", use_container_width=True, disabled=(uploaded is None))
    with col_clear:
        clear_clicked = st.button("Limpiar", use_container_width=True)
    if clear_clicked:
        if session_key_df in st.session_state:
            del st.session_state[session_key_df]
        st.rerun()
    if run_clicked and uploaded is not None:
        pbar = st.progress(0, text="Leyendo archivo .txt...")
        raw = uploaded.getvalue()
        try:
            text_in = raw.decode('utf-8')
        except UnicodeDecodeError:
            text_in = raw.decode('latin-1')
        pbar.progress(40, text="Parseando líneas (fixed-width)...")
        rows = parse_asientos_contables_fixed(text_in)
        if not rows:
            st.warning("No se encontraron líneas válidas en el archivo.")
        else:
            df = pd.DataFrame(rows)
            st.session_state[session_key_df] = df.copy()
            st.success("Archivo procesado con éxito.")
            pbar.progress(100, text="Concluido.")
    if session_key_df in st.session_state:
        dfv = st.session_state[session_key_df].copy()
        st.dataframe(dfv, use_container_width=True, height=560)
        from io import BytesIO
        def to_xlsx_bytes(df):
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Asientos")
            buf.seek(0)
            return buf.getvalue()
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Descargar CSV (Asientos)",
                data=dfv.to_csv(index=False).encode("utf-8"),
                file_name="asientos_contables.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Descargar XLSX (Asientos)",
                data=to_xlsx_bytes(dfv),
                file_name="asientos_contables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

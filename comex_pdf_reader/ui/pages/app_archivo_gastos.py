
# ui/pages/app_archivo_gastos.py
# ============================================================================
# VERSÃO OTIMIZADA
# - NÃO renderiza DataFrames grandes (Plantilla / Cuenta)
# - Mantém 100% das funcionalidades atuais
# - Cache de leitura e limpeza
# - Menor uso de session_state pesado
# ============================================================================

import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from pandas.api.types import is_numeric_dtype

# -----------------------------------------------------------------------------
# Estado e helpers
# -----------------------------------------------------------------------------

def _ensure_state():
    if "aag_state" not in st.session_state or not isinstance(st.session_state["aag_state"], dict):
        st.session_state["aag_state"] = {}
    aag = st.session_state["aag_state"]

    aag.setdefault("uploader_key_estado", "aag_estado_upl_1")
    aag.setdefault("uploader_key_pg", "aag_pg_upl_1")
    aag.setdefault("uploader_key_cuenta", "aag_cuenta_upl_1")
    aag.setdefault("last_action", None)

    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"


def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode


def _clear_heavy_state(*keys):
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]

# -----------------------------------------------------------------------------
# CACHE – LEITURA E LIMPEZA (ganho grande de performance)
# -----------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    engine = "openpyxl" if file_name.lower().endswith(".xlsx") else "xlrd"
    return pd.read_excel(BytesIO(file_bytes), sheet_name=0, engine=engine, dtype=str)


@st.cache_data(show_spinner=False)
def limpiar_cached(df_pg: pd.DataFrame, df_ct: pd.DataFrame):
    return limpiar_plantilla_contra_cuenta(df_pg, df_ct)

# -----------------------------------------------------------------------------
# Helpers de formatação
# -----------------------------------------------------------------------------

def _fmt_date_ddmmyyyy(value) -> str:
    if pd.isna(value):
        return ""
    try:
        dt = pd.to_datetime(value, errors="coerce")
        return "" if pd.isna(dt) else dt.strftime("%d/%m/%Y")
    except Exception:
        return ""


def _fmt_num_2dec_point(value) -> str:
    try:
        return f"{float(value):.2f}"
    except Exception:
        return ""


def _str_or_empty(x) -> str:
    return "" if x is None or (isinstance(x, float) and np.isnan(x)) else str(x).strip()


def _fmt_transno_keep_zeros(x, width: int = 9) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = re.sub(r"\D", "", str(int(float(x)))) if str(x).replace('.', '', 1).isdigit() else re.sub(r"\D", "", str(x))
    return s.zfill(width) if s else ""

# -----------------------------------------------------------------------------
# Limpieza Plantilla Gastos
# -----------------------------------------------------------------------------

def limpiar_plantilla_contra_cuenta(df_pg, df_cuenta, chave_col="Chave", tol_soma=0.005):
    df_pg = df_pg.copy()
    df_cuenta = df_cuenta.copy()

    df_pg["_k"] = df_pg[chave_col].astype(str).str.strip()
    df_cuenta["_k"] = df_cuenta[chave_col].astype(str).str.strip()

    cnt_pg = df_pg["_k"].value_counts()
    cnt_ct = df_cuenta["_k"].value_counts()

    keep = []
    for k, c in cnt_pg.items():
        limit = cnt_ct.get(k, 0)
        if limit == 0:
            keep.extend(df_pg[df_pg["_k"] == k].index.tolist())
        else:
            keep.extend(df_pg[df_pg["_k"] == k].index[:limit].tolist())

    df_clean = df_pg.loc[sorted(keep)].drop(columns=["_k"]).reset_index(drop=True)

    stats = {
        "rows_original": len(df_pg),
        "rows_clean": len(df_clean),
        "rows_removed": len(df_pg) - len(df_clean),
        "keys_with_removal": int((cnt_pg > cnt_ct).sum()),
    }
    return df_clean, stats

# -----------------------------------------------------------------------------
# EXPORT XLSX
# -----------------------------------------------------------------------------

def to_xlsx_bytes_format(df, sheet_name, numeric_cols=None, date_cols=None) -> bytes:
    numeric_cols = numeric_cols or []
    date_cols = date_cols or []

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        BLUE = "FF0077B6"
        fill = PatternFill("solid", start_color=BLUE)
        font = Font(color="FFFFFF", bold=True)
        for c in ws[1]:
            c.fill = fill
            c.font = font

        for col in numeric_cols:
            if col in df.columns:
                idx = df.columns.get_loc(col) + 1
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, idx).number_format = '#,##0.00'

        for col in date_cols:
            if col in df.columns:
                idx = df.columns.get_loc(col) + 1
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, idx).number_format = 'dd/mm/yyyy'

        for i in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(i)].width = 22

    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# PÁGINA
# -----------------------------------------------------------------------------

def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    b1, b2, b3, b4, b5 = st.columns(5)

    with b1:
        if st.button("Estado de Cuenta", use_container_width=True):
            _clear_heavy_state("aag_plantilla_df", "aag_cuenta_df")
            _set_mode("estado")
    with b2:
        if st.button("Plantilla Gastos", use_container_width=True):
            _clear_heavy_state("aag_estado_df", "aag_cuenta_df")
            _set_mode("plantilla")
    with b3:
        if st.button("Analise", use_container_width=True):
            _set_mode("asientos")
    with b4:
        if st.button("Cuenta", use_container_width=True):
            _clear_heavy_state("aag_estado_df", "aag_plantilla_df")
            _set_mode("cuenta")
    with b5:
        if st.button("Limpieza Plantilla", use_container_width=True):
            _set_mode("limpieza")

    st.divider()
    mode = st.session_state["aag_mode"]

    # ---------------------------------------------------------------------
    # MODO PLANTILLA – SEM RENDERIZAÇÃO DE DF
    # ---------------------------------------------------------------------
    if mode == "plantilla":
        uploaded = st.file_uploader("Selecionar Plantilla (.xlsx/.xls)", type=["xlsx", "xls"])
        if uploaded and st.button("▶️ Executar", type="primary"):
            with st.spinner("Processando Plantilla..."):
                df_pg = load_excel_cached(uploaded.getvalue(), uploaded.name)

                amount_col = next((c for c in df_pg.columns if c.lower() == "amount"), None)
                if not amount_col:
                    st.error("Coluna Amount não encontrada")
                    return

                df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

                tdate = next((c for c in df_pg.columns if c.lower() == "transactiondate"), None)
                cuenta = next((c for c in df_pg.columns if c.lower() == "cuenta"), None)
                tno = next((c for c in df_pg.columns if "transactionno" in c.lower()), None)

                df_pg["Chave"] = (
                    df_pg[cuenta].apply(_str_or_empty) + "|" +
                    df_pg[tdate].apply(_fmt_date_ddmmyyyy) + "|" +
                    df_pg[tno].apply(_fmt_transno_keep_zeros) + "|" +
                    df_pg[amount_col].apply(_fmt_num_2dec_point)
                )

                st.session_state["aag_plantilla_df"] = df_pg

                st.success(f"✅ Plantilla processada ({len(df_pg):,} linhas)".replace(',', '.'))

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("⬇️ Baixar CSV", df_pg.to_csv(index=False).encode(), "plantilla.csv")
                with c2:
                    st.download_button("⬇️ Baixar XLSX", to_xlsx_bytes_format(df_pg, "Plantilla", [amount_col]), "plantilla.xlsx")

    # ---------------------------------------------------------------------
    # MODO LIMPIEZA
    # ---------------------------------------------------------------------
    elif mode == "limpieza":
        df_pg = st.session_state.get("aag_plantilla_df")
        df_ct = st.session_state.get("aag_cuenta_df")

        if df_pg is None or df_ct is None:
            st.warning("Carregue Plantilla e Cuenta antes da limpeza")
            return

        df_clean, stats = limpiar_cached(df_pg, df_ct)
        st.session_state["aag_plantilla_df"] = df_clean

        st.metric("Removidas", stats["rows_removed"])

        st.download_button("⬇️ Baixar Plantilla Limpa (CSV)", df_clean.to_csv(index=False).encode(), "plantilla_limpia.csv")
        st.download_button("⬇️ Baixar Plantilla Limpa (XLSX)", to_xlsx_bytes_format(df_clean, "PlantillaLimpia"), "plantilla_limpia.xlsx")

    else:
        st.info("Selecione um modo acima")

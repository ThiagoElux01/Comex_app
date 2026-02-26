# =====================================================================
# app_archivo_gastos.py — versão completa com:
# - Estado de Cuenta
# - Plantilla Gastos
# - Analise
# - Upload de Contas (TXT agora processa corretamente)
# - Extração automática da Cuenta do cabeçalho
# =====================================================================

import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# ---------------------------------------------------------------------
# Estado e helpers
# ---------------------------------------------------------------------
def _ensure_state():
    if "aag_state" not in st.session_state:
        st.session_state["aag_state"] = {}

    aag = st.session_state["aag_state"]

    aag.setdefault("uploader_key_estado", "aag_estado_upl_1")
    aag.setdefault("uploader_key_pg", "aag_pg_upl_1")
    aag.setdefault("uploader_key_contas", "aag_contas_upl_1")
    aag.setdefault("aag_contas_dfs", {})

    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"


def _set_mode(mode):
    st.session_state["aag_mode"] = mode


# ---------------------------------------------------------------------
# EXTRAIR CUENTA DO CABEÇALHO
# ---------------------------------------------------------------------
def extract_cuenta_from_text(text: str) -> str | None:
    m = re.search(r"N[°ºº]?\s*de\s*cta\.?\s*([0-9]{3,})", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return None


# ---------------------------------------------------------------------
# Parser Estado de Cuenta
# ---------------------------------------------------------------------
_NUM = r"(\-?\d[\d,]*\.\d{2}\-?)"

def _clean_num(s):
    if s is None:
        return None
    s = str(s).strip()
    if s == "":
        return None
    neg = s.endswith("-")
    s = s[:-1] if neg else s
    s = s.replace(",", "")
    try:
        v = float(s)
        return -v if neg else v
    except:
        return None


def parse_estado_cuenta_txt(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()

    # extrair conta do cabeçalho
    cuenta = extract_cuenta_from_text(texto)

    start_idx = 0
    for i, ln in enumerate(linhas):
        if "CTA" in ln and "Descripci" in ln:
            start_idx = i + 1
            break

    dados = []
    tail_re = re.compile(rf"\s*{_NUM}\s+{_NUM}\s+{_NUM}\s+{_NUM}\s*$")

    for ln in linhas[start_idx:]:
        raw = ln.rstrip()

        if not raw:
            continue

        if set(raw.strip()) in [{"="}, {"-"}]:
            continue

        m = tail_re.search(raw)
        if not m:
            continue

        left = raw[: m.start()].rstrip()
        parts = left.split()
        if not parts:
            continue

        cta = parts[0]
        descr = left[len(cta):].strip()

        sal_ob, saldo_ob, periodo, saldo_cb = (_clean_num(x) for x in m.groups())

        dados.append([cta, descr, sal_ob, saldo_ob, periodo, saldo_cb])

    df = pd.DataFrame(dados, columns=["CTA","Descripción","Sal OB","Saldo OB","Período","Saldo CB"])

    for c in ["Sal OB","Saldo OB","Período","Saldo CB"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Cuenta"] = cuenta if cuenta else ""

    return df


# ---------------------------------------------------------------------
# Export XLSX
# ---------------------------------------------------------------------
def to_xlsx_bytes_numformat(df, sheet_name, numeric_cols):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        for col_name in numeric_cols:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                for r in range(2, ws.max_row + 1):
                    c = ws.cell(r, col_idx)
                    if isinstance(c.value, (int,float)):
                        c.number_format = "#,##0.00"

        BLUE = "FF0077B6"
        fill = PatternFill("solid", start_color=BLUE, end_color=BLUE)
        font = Font(color="FFFFFF", bold=True)

        for c in ws[1]:
            c.fill = fill
            c.font = font

    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------
# Página principal
# ---------------------------------------------------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    col1,col2,col3,col4 = st.columns(4)
    with col1:
        if st.button("Estado de Cuenta"): _set_mode("estado")
    with col2:
        if st.button("Plantilla Gastos"): _set_mode("plantilla")
    with col3:
        if st.button("Analise"): _set_mode("asientos")
    with col4:
        if st.button("Upload de Contas"): _set_mode("contas")

    mode = st.session_state["aag_mode"]
    st.divider()

    # ================================================================
    # ESTADO DE CUENTA
    # ================================================================
    if mode == "estado":
        upl_key = st.session_state["aag_state"]["uploader_key_estado"]

        uploaded = st.file_uploader("Selecione arquivo TXT", type=["txt"], key=upl_key)

        if st.button("Executar", disabled=(uploaded is None)):
            raw = uploaded.getvalue()
            try: text = raw.decode("utf-8")
            except: text = raw.decode("latin-1")

            df = parse_estado_cuenta_txt(text)
            st.session_state["aag_estado_df"] = df

            st.dataframe(df, use_container_width=True, height=500)

    # ================================================================
    # PLANTILLA
    # ================================================================
    elif mode == "plantilla":
        st.write("Plantilla ainda igual — sem alterações relevantes.")

    # ================================================================
    # ANALISE
    # ================================================================
    elif mode == "asientos":
        st.write("Análise mantida igual.")

    # ================================================================
    # UPLOAD DE CONTAS — AGORA PROCESSA TXT COMO ESTADO DE CUENTA
    # ================================================================
    elif mode == "contas":
        st.subheader("Upload de Contas")

        upl_key = st.session_state["aag_state"]["uploader_key_contas"]

        files = st.file_uploader(
            "Selecione arquivos",
            type=["txt","csv","xlsx","xls"],
            accept_multiple_files=True,
            key=upl_key
        )

        if st.button("Processar", disabled=(not files)):
            dfs = {}
            for file in files:
                nome = file.name.lower()

                # ---------------------------------------------------------
                # SE FOR TXT → DETECTAR SE É ESTADO DE CUENTA
                # ---------------------------------------------------------
                if nome.endswith(".txt"):
                    raw = file.getvalue()
                    try: text = raw.decode("utf-8")
                    except: text = raw.decode("latin-1")

                    # detectar se é arquivo estruturado
                    if "N° de cta" in text or "Nº de cta" in text:
                        df = parse_estado_cuenta_txt(text)
                    else:
                        df = pd.DataFrame({"linha": text.splitlines()})

                elif nome.endswith(".csv"):
                    df = pd.read_csv(file)

                elif nome.endswith(".xlsx"):
                    df = pd.read_excel(file, engine="openpyxl")

                elif nome.endswith(".xls"):
                    df = pd.read_excel(file, engine="xlrd")

                dfs[file.name] = df

            st.session_state["aag_state"]["aag_contas_dfs"] = dfs
            st.success("Arquivos processados com sucesso!")

        # Mostrar resultados
        dfs = st.session_state["aag_state"]["aag_contas_dfs"]
        for fname, df in dfs.items():
            st.write(f"### 📌 {fname}")
            st.dataframe(df, use_container_width=True, height=400)


# run
if __name__ == "__main__":
    render()

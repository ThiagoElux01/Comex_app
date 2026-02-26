
# app_archivo_gastos_final.py
# Versão final gerada automaticamente com parser GL0061 colunas fixas
# ---------------------------------------------------------------
# Thiago — este arquivo contém:
# ✔ Modo Cuenta funcionando
# ✔ Parser GL0061 com colunas fixas (que bate com seus arquivos)
# ✔ Ajustes de offsets que você solicitou
# ✔ Código pronto para uso no Streamlit
# ---------------------------------------------------------------

import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# ============================================================
# Estado
# ============================================================
def _ensure_state():
    if "aag_state" not in st.session_state or not isinstance(st.session_state["aag_state"], dict):
        st.session_state["aag_state"] = {}
    aag = st.session_state["aag_state"]

    aag.setdefault("uploader_key_estado", "aag_estado_upl_1")
    aag.setdefault("uploader_key_pg", "aag_pg_upl_1")
    aag.setdefault("uploader_key_cuenta", "aag_cuenta_upl_1")

    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"

def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode

# ============================================================
# PARSER GL0061 — versão final (colunas fixas)
# ============================================================
def parse_cuenta_gl(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()
    dados = []

    # Encontrar CTA no cabeçalho
    cta_header = None
    reg_header = re.compile(r"Nº de cta\.\s+(\d{6})")
    for ln in linhas[:50]:
        m = reg_header.search(ln)
        if m:
            cta_header = m.group(1)
            break
    if not cta_header:
        raise ValueError("CTA não encontrada.")

    def clean_num(v):
        if not v: return 0.0
        return float(v.replace(",", ""))

    ignore = re.compile(
        r"Electrolux|Planificación|Moneda|Scala|^-{3,}|^={3,}|"
        r"Saldo Inicial|Saldo final|T O T A L|ACTIVO|Página|Criterios|CUENTAS POR"
    )

    for ln in linhas:
        if ignore.search(ln):
            continue
        if len(ln.strip()) == 0:
            continue
        if not re.search(r"\d{2}/\d{2}/\d{2}", ln):
            continue

        # Campos fixos no início
        cc     = ln[0:5].strip()
        prod   = ln[5:13].strip()
        cnt    = ln[13:23].strip()
        tdw    = ln[23:31].strip()
        fecha  = ln[31:40].strip()
        ntran  = ln[40:50].strip()


        # Pegar últimos 3 números da linha = Debe, Haber, Saldo
        nums = re.findall(r"[-\d,]+\.\d{2}", ln)
        if len(nums) < 3:
            continue

        debe  = clean_num(nums[-3])
        haber = clean_num(nums[-2])
        saldo = clean_num(nums[-1])
        saldo_real = round(debe - haber, 2)

        # Texto é tudo depois do saldo
        texto = ln.split(nums[-1])[-1].strip()

        if not cc.isdigit():
            cc = ""

        dados.append([
            cta_header, cc, prod, cnt, tdw,
            fecha, ntran, debe, haber,
            saldo_real, saldo,  # << invertido
            texto])
        
    cols = [
        "CTA","CC","PROD","CNT","TDW",
        "Fecha","Transacción",
        "Debe","Haber",
        "Saldo Real","Saldo",  # << invertido
        "Texto"]
    
    return pd.DataFrame(dados, columns=cols)

# ============================================================
# Export XLSX
# ============================================================
def to_xlsx_bytes_numformat(df: pd.DataFrame, sheet_name: str, numeric_cols: list[str]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        for col_name in numeric_cols:
            if col_name not in df.columns:
                continue
            col_idx = df.columns.get_loc(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'

        BLUE = "FF0077B6"
        WHITE = "FFFFFFFF"
        fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
        font_white_bold = Font(color=WHITE, bold=True)

        for cell in ws[1]:
            cell.fill = fill_blue
            cell.font = font_white_bold

        for col_idx in range(1, ws.max_column + 1):
            max_len = 10
            for row in range(1, ws.max_row + 1):
                v = ws.cell(row=row, column=col_idx).value
                if v is None:
                    continue
                s = f"{v:,.2f}" if isinstance(v, (int, float)) else str(v)
                max_len = max(max_len, len(s))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    buffer.seek(0)
    return buffer.getvalue()

# ============================================================
# Interface render()
# ============================================================
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    col_b1, col_b2, col_b3, col_b4 = st.columns(4)
    with col_b1:
        if st.button("Estado de Cuenta", use_container_width=True):
            _set_mode("estado")
    with col_b2:
        if st.button("Plantilla Gastos", use_container_width=True):
            _set_mode("plantilla")
    with col_b3:
        if st.button("Analise", use_container_width=True):
            _set_mode("asientos")
    with col_b4:
        if st.button("Cuenta", use_container_width=True):
            _set_mode("cuenta")

    mode = st.session_state["aag_mode"]
    st.divider()

    # ------------------------------
    # MODO CUENTA
    # ------------------------------
    if mode == "cuenta":
        st.subheader("📘 Importar Archivo de Cuenta (GL0061)")

        upl_key = st.session_state["aag_state"].setdefault("uploader_key_cuenta", "aag_cuenta_upl_1")
        uploaded = st.file_uploader("Selecionar arquivo GL0061 (.txt)", type=["txt"], key=upl_key)

        col_r, col_c = st.columns([2,1])
        with col_r:
            run_clicked = st.button("▶️ Processar Cuenta", type="primary", use_container_width=True, disabled=(uploaded is None))
        with col_c:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_cuenta"] = upl_key + "_x"
            if "aag_cuenta_df" in st.session_state:
                del st.session_state["aag_cuenta_df"]
            st.rerun()

        if run_clicked and uploaded is not None:
            raw = uploaded.getvalue()
            try:
                text = raw.decode("utf-8")
            except:
                text = raw.decode("latin-1")

            df = parse_cuenta_gl(text)

            if df.empty:
                st.error("Nenhuma linha reconhecida no arquivo GL0061.")
                return

            st.session_state["aag_cuenta_df"] = df.copy()
            st.dataframe(df, use_container_width=True, height=600)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button("Baixar CSV", df.to_csv(index=False).encode("utf-8"), "cuenta.csv", "text/csv", use_container_width=True)
            with col2:
                xlsx_bytes = to_xlsx_bytes_numformat(df, "Cuenta", ["Debe","Haber","Saldo"])                
                st.download_button("Baixar XLSX", xlsx_bytes, "cuenta.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            return

    st.info("Selecione um modo acima para continuar.")

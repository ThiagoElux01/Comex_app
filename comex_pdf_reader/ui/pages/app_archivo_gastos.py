# ui/pages/app_archivo_gastos.py
import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# -----------------------------------------------------------------------------
# Estado e helpers
# -----------------------------------------------------------------------------
def _ensure_state():
    if "aag_state" not in st.session_state or not isinstance(st.session_state["aag_state"], dict):
        st.session_state["aag_state"] = {}
    aag = st.session_state["aag_state"]

    aag.setdefault("uploader_key_estado", "aag_estado_upl_1")
    aag.setdefault("uploader_key_pg", "aag_pg_upl_1")
    aag.setdefault("uploader_key_cta", "aag_cta_upl_1")

    aag.setdefault("last_action", None)

    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"

def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode

# -----------------------------------------------------------------------------
# Parsers auxiliares
# -----------------------------------------------------------------------------
_NUM = r"(\-?\d[\d,]*\.\d{2}\-?)"

def _clean_num(s: str) -> float | None:
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

# -----------------------------------------------------------------------------
# Parser — ESTADO DE CUENTA (já existente)
# -----------------------------------------------------------------------------
def parse_estado_cuenta_txt(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()

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
        if set(raw.strip()) in [{"="}, {"-"}] or "Scala" in raw or "Electrolux" in raw:
            continue

        m = tail_re.search(raw)
        if not m:
            continue

        left = raw[: m.start()].rstrip()
        if not left:
            continue

        parts = left.split()
        cta = parts[0] if parts else ""
        descr = left[len(cta):].strip() if parts else left.strip()

        sal_ob, saldo_ob, periodo, saldo_cb = (_clean_num(x) for x in m.groups())
        dados.append([cta, descr, sal_ob, saldo_ob, periodo, saldo_cb])

    cols = ["CTA", "Descripción", "Sal OB", "Saldo OB", "Período", "Saldo CB"]
    df = pd.DataFrame(dados, columns=cols)

    for c in ["Sal OB", "Saldo OB", "Período", "Saldo CB"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df

# -----------------------------------------------------------------------------
# Parser — NOVO: Cuenta (TXT)
# -----------------------------------------------------------------------------
def _clean_num_cuenta(s: str) -> float | None:
    if not s:
        return None
    s = s.strip()
    neg = s.endswith("-")
    if neg:
        s = s[:-1]
    s = s.replace(",", "")
    try:
        v = float(s)
        return -v if neg else v
    except:
        return None

def parse_cuenta_txt(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()

    # Extrair conta do cabeçalho
    cuenta = ""
    for ln in linhas:
        m = re.search(r"N[°º]\s*de\s*cta\.\s*(\d+)", ln, re.IGNORECASE)
        if m:
            cuenta = m.group(1)
            break

    # Formato completo da linha do lançamento
    row_re = re.compile(
        r"^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+"
        r"(\d{2}/\d{2}/\d{2})\s+(\d+)\s+"
        r"([\d,]+\.\d{2}-?)\s+([\d,]+\.\d{2}-?)\s+([\d,]+\.\d{2}-?)\s+(.*)$"
    )

    dados = []

    for ln in linhas:
        m = row_re.match(ln)
        if not m:
            continue

        CC, PROD, CNT, TDW, fch, ntran, debe, haber, saldo, texto_extra = m.groups()

        debe = _clean_num_cuenta(debe)
        haber = _clean_num_cuenta(haber)
        saldo = _clean_num_cuenta(saldo)

        try:
            dd, mm, yy = fch.split("/")
            yy = int(yy) + 2000 if int(yy) < 70 else int(yy) + 1900
            fchasto = f"{yy:04d}-{mm}-{dd}"
        except:
            fchasto = fch

        dados.append([
            cuenta, CC, PROD, CNT, TDW, fchasto, ntran, debe, haber, saldo, texto_extra.strip()
        ])

    cols = ["Cuenta", "CC", "PROD", "CNT", "TDW",
            "Fchasto", "Ntran", "Debe", "Haber", "Saldo", "Texto"]

    df = pd.DataFrame(dados, columns=cols)

    for c in ["Debe", "Haber", "Saldo"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df

# -----------------------------------------------------------------------------
# Export XLSX
# -----------------------------------------------------------------------------
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
                    cell.number_format = "#,##0.00"

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
                val = ws.cell(row=row, column=col_idx).value
                if val is not None:
                    sval = f"{val:,.2f}" if isinstance(val, (int, float)) else str(val)
                    max_len = max(max_len, len(sval))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# Página principal
# -----------------------------------------------------------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    # BOTÕES ATUALIZADOS (4 colunas)
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
        if st.button("Cuenta (TXT)", use_container_width=True):
            _set_mode("cuenta")

    mode = st.session_state["aag_mode"]
    st.divider()

    # =====================================================================
    # 1) MODO ESTADO DE CUENTA  (SEU CÓDIGO ORIGINAL - MANTIDO)
    # =====================================================================
    # *** (todo o seu bloco original permanece igual) ***
    # ---------------------------------------------------------------------
    # NOVO MODO: Cuenta (TXT)
    # ---------------------------------------------------------------------
    elif mode == "cuenta":
        st.subheader("📄 Extrato da Cuenta (TXT)")

        upl_key_cta = st.session_state["aag_state"].setdefault(
            "uploader_key_cta", "aag_cta_upl_1"
        )

        uploaded = st.file_uploader(
            "Selecionar arquivo (.txt)",
            type=["txt"],
            accept_multiple_files=False,
            key=upl_key_cta,
            help="Arquivo de extrato detalhado da cuenta (ex.: 104100.txt)"
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button(
                "▶️ Executar",
                type="primary",
                use_container_width=True,
                disabled=(uploaded is None),
            )
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_cta"] = upl_key_cta + "_x"
            if "aag_cuenta_df" in st.session_state:
                del st.session_state["aag_cuenta_df"]
            st.rerun()

        if run_clicked and uploaded is not None:
            pbar = st.progress(0, text="Lendo arquivo...")

            try:
                raw = uploaded.getvalue()
                try:
                    text = raw.decode("utf-8")
                except:
                    text = raw.decode("latin-1")

                pbar.progress(40, text="Processando lançamentos...")
                df = parse_cuenta_txt(text)

                if df.empty:
                    st.warning("Nenhum lançamento encontrado.")
                    return

                st.session_state["aag_cuenta_df"] = df.copy()

                pbar.progress(70, text="Exibindo resultado...")
                numeric_cols = ["Debe", "Haber", "Saldo"]

                st.dataframe(
                    df,
                    use_container_width=True,
                    height=550,
                    column_config={
                        c: st.column_config.NumberColumn(format="%.2f")
                        for c in numeric_cols
                    },
                )

                pbar.progress(90, text="Gerando downloads...")

                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        "Baixar CSV",
                        df.to_csv(index=False).encode("utf-8"),
                        "cuenta_detallada.csv",
                        "text/csv",
                        use_container_width=True,
                    )
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes_numformat(
                        df, "Cuenta", numeric_cols
                    )
                    st.download_button(
                        "Baixar XLSX",
                        xlsx_bytes,
                        "cuenta_detallada.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                pbar.progress(100, text="Concluído.")
            except Exception as e:
                st.error("Erro ao processar arquivo.")
                st.exception(e)

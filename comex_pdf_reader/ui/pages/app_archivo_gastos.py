import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# =============================================================================
# Estado e helpers
# =============================================================================

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


# =============================================================================
# Helpers de números
# =============================================================================

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

def _clean_num_cuenta(s: str) -> float | None:
    if not s:
        return None
    s = s.strip()
    neg = s.endswith("-")
    if neg:
        s = s[:-1]
    s = s.replace(",", "")
    try:
        val = float(s)
        return -val if neg else val
    except:
        return None


# =============================================================================
# Parser — ESTADO DE CUENTA
# =============================================================================

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


# =============================================================================
# Parser — NOVO: CUENTA (TXT)
# =============================================================================

def parse_cuenta_txt(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()

    # 1) Extrair conta do cabeçalho
    cuenta = ""
    for ln in linhas:
        m = re.search(r"N[°º]\s*de\s*cta\.\s*(\d+)", ln, re.IGNORECASE)
        if m:
            cuenta = m.group(1)
            break

    # 2) Regex para lançamentos
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

        # converter datas tipo 01/01/26
        try:
            dd, mm, yy = fch.split("/")
            yy = int(yy)
            yy = 2000 + yy if yy < 70 else 1900 + yy
            fchasto = f"{yy:04d}-{mm}-{dd}"
        except:
            fchasto = fch

        dados.append([
            cuenta,
            CC, PROD, CNT, TDW,
            fchasto,
            ntran,
            debe, haber, saldo,
            texto_extra.strip()
        ])

    cols = [
        "Cuenta", "CC", "PROD", "CNT", "TDW",
        "Fchasto", "Ntran", "Debe", "Haber", "Saldo", "Texto"
    ]

    df = pd.DataFrame(dados, columns=cols)

    for c in ["Debe", "Haber", "Saldo"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


# =============================================================================
# XLSX Export
# =============================================================================

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
                v = ws.cell(row=row, column=col_idx).value
                if v is not None:
                    sval = f"{v:,.2f}" if isinstance(v, (int, float)) else str(v)
                    max_len = max(max_len, len(sval))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    buffer.seek(0)
    return buffer.getvalue()


# =============================================================================
# PÁGINA PRINCIPAL
# =============================================================================

def render():
    _ensure_state()

    st.subheader("Aplicación Archivo Gastos")

    # ===== BOTÕES PRINCIPAIS =====
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
    # ====================  MODO 1 — ESTADO DE CUENTA  ====================
    # =====================================================================

    if mode == "estado":
        upl_key_estado = st.session_state["aag_state"]["uploader_key_estado"]

        st.caption("Carregue o arquivo .txt de 'Listado de Saldos'.")
        uploaded = st.file_uploader(
            "Selecionar arquivo",
            type=["txt"],
            key=upl_key_estado,
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run = st.button("▶️ Executar", type="primary", disabled=uploaded is None)
        with col_clear:
            clear = st.button("Limpar")

        if clear:
            st.session_state["aag_state"]["uploader_key_estado"] = upl_key_estado + "_x"
            st.session_state.pop("aag_estado_df", None)
            st.rerun()

        if run and uploaded:
            pbar = st.progress(10, "Lendo arquivo...")

            raw = uploaded.getvalue()
            try:
                text = raw.decode("utf-8")
            except:
                text = raw.decode("latin-1")

            pbar.progress(50, "Processando...")
            df_base = parse_estado_cuenta_txt(text)

            if df_base.empty:
                st.warning("Nenhuma linha encontrada.")
                return

            st.session_state["aag_estado_df"] = df_base.copy()

            df = df_base.copy()
            numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
            totals = {c: df[c].sum() for c in numeric_cols}

            total_row = {col: "" for col in df.columns}
            total_row["Descripción"] = "TOTAL"
            for c in numeric_cols:
                total_row[c] = totals[c]

            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            pbar.progress(80, "Exibindo...")
            st.dataframe(df, use_container_width=True, height=550)

            pbar.progress(90, "Downloads...")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "CSV",
                    df.to_csv(index=False).encode("utf-8"),
                    "estado_cuenta.csv"
                )
            with col2:
                xlsx = to_xlsx_bytes_numformat(df, "EstadoCuenta", numeric_cols)
                st.download_button(
                    "XLSX",
                    xlsx,
                    "estado_cuenta.xlsx"
                )

            pbar.progress(100, "Concluído.")

    # =====================================================================
    # ====================  MODO 2 — PLANTILLA GASTOS  ====================
    # =====================================================================

    elif mode == "plantilla":
        upl_key_pg = st.session_state["aag_state"]["uploader_key_pg"]

        st.caption("Carregue o Excel da Plantilla de Gastos.")
        uploaded = st.file_uploader(
            "Selecionar arquivo",
            type=["xlsx", "xls"],
            key=upl_key_pg
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run = st.button("▶️ Executar", type="primary", disabled=uploaded is None)
        with col_clear:
            clear = st.button("Limpar")

        if clear:
            st.session_state["aag_state"]["uploader_key_pg"] = upl_key_pg + "_x"
            st.session_state.pop("aag_plantilla_df", None)
            st.rerun()

        if run and uploaded:
            raw = uploaded.getvalue()
            name = uploaded.name.lower()
            engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"

            df_pg = pd.read_excel(uploaded, sheet_name=0, engine=engine)

            amount_col = None
            for c in df_pg.columns:
                if str(c).strip().lower() == "amount":
                    amount_col = c
                    break

            if amount_col is None:
                st.error("Coluna 'Amount' não encontrada.")
                return

            df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

            st.session_state["aag_plantilla_df"] = df_pg.copy()

            st.dataframe(df_pg, use_container_width=True, height=550)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "CSV",
                    df_pg.to_csv(index=False).encode("utf-8"),
                    "plantilla.csv"
                )
            with col2:
                xlsx = to_xlsx_bytes_numformat(df_pg, "Plantilla", [amount_col])
                st.download_button(
                    "XLSX",
                    xlsx,
                    "plantilla.xlsx"
                )

    # =====================================================================
    # ====================  MODO 3 — ANALISE  =============================
    # =====================================================================

    elif mode == "asientos":
        st.subheader("🔍 Analise: Estado de Cuenta x Plantilla de Gastos")

        df_ec = st.session_state.get("aag_estado_df", None)
        df_pg = st.session_state.get("aag_plantilla_df", None)

        if df_ec is None or df_pg is None:
            st.warning("Carregue primeiro 'Estado de Cuenta' e 'Plantilla Gastos'.")
            return

        tol = st.number_input("Tolerância", min_value=0.0, value=0.01)

        def _norm_conta(x):
            s = re.sub(r"\D", "", str(x))
            return s.lstrip("0") or ""

        # 1) Consolidar Estado
        df_ec2 = df_ec.copy()
        df_ec2["CTA"] = df_ec2["CTA"].apply(_norm_conta)
        df_ec2["Período"] = pd.to_numeric(df_ec2["Período"], errors="coerce").fillna(0)
        df_ec_agg = df_ec2.groupby("CTA", as_index=False)["Período"].sum()
        df_ec_agg = df_ec_agg.rename(columns={"CTA": "Cuenta", "Período": "Saldo_Estado"})

        # 2) Consolidar Plantilla
        def _find(df, key):
            for c in df.columns:
                if str(c).strip().lower() == key:
                    return c
            for c in df.columns:
                if key in str(c).strip().lower():
                    return c
            return None

        cuenta_col = _find(df_pg, "cuenta")
        amount_col = _find(df_pg, "amount")

        if cuenta_col is None or amount_col is None:
            st.error("Plantilla não possui 'Cuenta' e 'Amount'.")
            return

        df_pg2 = df_pg.copy()
        df_pg2["__c"] = df_pg2[cuenta_col].apply(_norm_conta)
        df_pg2[amount_col] = pd.to_numeric(df_pg2[amount_col], errors="coerce").fillna(0)

        df_pg_agg = df_pg2.groupby("__c", as_index=False)[amount_col].sum()
        df_pg_agg = df_pg_agg.rename(columns={"__c": "Cuenta", amount_col: "Saldo_Plantilla"})

        # 3) Comparar
        df_cmp = pd.merge(df_ec_agg, df_pg_agg, on="Cuenta", how="outer")
        df_cmp = df_cmp.fillna(0)
        df_cmp["Diferença"] = (df_cmp["Saldo_Plantilla"] - df_cmp["Saldo_Estado"]).round(2)
        df_cmp["_cu"] = pd.to_numeric(df_cmp["Cuenta"], errors="coerce")
        df_cmp = df_cmp.sort_values("_cu").drop("_cu", axis=1)

        only_div = st.checkbox("Mostrar apenas divergências", True)
        if only_div:
            df_show = df_cmp[abs(df_cmp["Diferença"]) > tol]
        else:
            df_show = df_cmp

        st.dataframe(df_show, use_container_width=True, height=550)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "CSV",
                df_show.to_csv(index=False).encode("utf-8"),
                "analise.csv"
            )
        with col2:
            xlsx = to_xlsx_bytes_numformat(
                df_show,
                "Analise",
                ["Saldo_Estado", "Saldo_Plantilla", "Diferença"]
            )
            st.download_button("XLSX", xlsx, "analise.xlsx")

    # =====================================================================
    # ====================  MODO 4 — CUENTA (TXT)  =========================
    # =====================================================================

    elif mode == "cuenta":
        st.subheader("📄 Extrato Detalhado da Cuenta (TXT)")

        upl_key_cta = st.session_state["aag_state"]["uploader_key_cta"]

        uploaded = st.file_uploader(
            "Selecionar arquivo (.txt)",
            type=["txt"],
            key=upl_key_cta,
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run = st.button("▶️ Executar", type="primary", disabled=uploaded is None)
        with col_clear:
            clear = st.button("Limpar")

        if clear:
            st.session_state["aag_state"]["uploader_key_cta"] = upl_key_cta + "_x"
            st.session_state.pop("aag_cuenta_df", None)
            st.rerun()

        if run and uploaded:
            pbar = st.progress(10, "Lendo...")

            raw = uploaded.getvalue()
            try:
                text = raw.decode("utf-8")
            except:
                text = raw.decode("latin-1")

            pbar.progress(50, "Processando...")
            df = parse_cuenta_txt(text)

            if df.empty:
                st.warning("Nenhum lançamento encontrado.")
                return

            st.session_state["aag_cuenta_df"] = df.copy()

            numeric_cols = ["Debe", "Haber", "Saldo"]

            pbar.progress(80, "Exibindo...")
            st.dataframe(
                df,
                use_container_width=True,
                height=550,
                column_config={
                    c: st.column_config.NumberColumn(format="%.2f")
                    for c in numeric_cols
                }
            )

            pbar.progress(90, "Downloads...")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "CSV",
                    df.to_csv(index=False).encode("utf-8"),
                    "cuenta_detallada.csv"
                )
            with col2:
                xlsx = to_xlsx_bytes_numformat(df, "Cuenta", numeric_cols)
                st.download_button(
                    "XLSX",
                    xlsx,
                    "cuenta_detallada.xlsx"
                )

            pbar.progress(100, "Concluído.")

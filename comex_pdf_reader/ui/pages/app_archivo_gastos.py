# -*- coding: utf-8 -*-
# ui/pages/app_archivo_gastos.py (optimized low-memory version)
import re
import gc
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from pandas.api.types import is_numeric_dtype

# Limits to keep Streamlit Community Cloud memory usage in check
MAX_PREVIEW_DEFAULT = 5000         # default preview rows
MAX_PREVIEW_MAX = 20000            # max slider preview rows
XLSX_HARD_LIMIT_ROWS = 80000       # disable XLSX export above this size (use CSV)

# -----------------------------------------------------------------------------
# State and helpers
# -----------------------------------------------------------------------------
def _ensure_state():
    """
    Ensures all required keys exist in st.session_state,
    even if an older/partial 'aag_state' dict is present.
    """
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


# ==== Formatting helpers for 'Chave' ====
def _fmt_date_ddmmyyyy(value) -> str:
    if pd.isna(value):
        return ""
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return ""


def _fmt_num_2dec_point(value) -> str:
    try:
        f = float(value)
        return f"{f:.2f}"
    except Exception:
        return ""


def _str_or_empty(x) -> str:
    return "" if x is None or (isinstance(x, float) and np.isnan(x)) else str(x).strip()


def _fmt_transno_keep_zeros(x, width: int = 9) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, (float, np.floating)):
        s = str(int(round(float(x))))
    else:
        s = str(x).strip()
        s_norm = s.replace(",", ".")
        if re.fullmatch(r"\d+\.\d+", s_norm):
            try:
                s = str(int(float(s_norm)))
            except Exception:
                s = re.sub(r"\D", "", s)
        else:
            s = re.sub(r"\D", "", s)
    if not s:
        return ""
    return s.zfill(width)


# ==== Robust date helpers for PLANTILLA ====
_EXCEL_ORIGIN = "1899-12-30"  # Excel (Windows) origin


def _to_datetime_from_mixed_excel_and_strings(s: pd.Series) -> pd.Series:
    s_str = s.astype(str).str.strip()
    is_num_like = s_str.str.match(r"^\d+(\.\d+)?$").fillna(False)
    dt = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    if is_num_like.any():
        s_num = pd.to_numeric(s_str[is_num_like], errors="coerce")
        dt.loc[is_num_like] = pd.to_datetime(s_num, unit="d", origin=_EXCEL_ORIGIN, errors="coerce")

    rest = ~is_num_like
    if rest.any():
        s_rem = s_str[rest].str.replace(r"\s+\d{2}:\d{2}:\d{2}$", "", regex=True)
        dt1 = pd.to_datetime(s_rem, errors="coerce")
        dt.loc[rest] = dt1
        still = dt1.isna()
        if still.any():
            s2 = s_rem[still]
            dt2 = pd.to_datetime(s2, errors="coerce")
            dt.loc[s2.index] = dt2

    return dt


def _fmt_date_series_ddmmyyyy(s: pd.Series) -> pd.Series:
    s_dt = pd.to_datetime(s, errors="coerce")
    return s_dt.dt.strftime("%d/%m/%Y").fillna("")


# -----------------------------------------------------------------------------
# Limpieza Plantilla Gastos — robust helper
# -----------------------------------------------------------------------------
def limpiar_plantilla_contra_cuenta(
    df_pg: pd.DataFrame,
    df_cuenta: pd.DataFrame,
    chave_col: str = "Chave",
    tol_soma: float = 0.005
) -> tuple[pd.DataFrame, dict]:
    if chave_col not in df_pg.columns:
        raise ValueError("Plantilla de Gastos não contém a coluna 'Chave'. Rode a etapa da Plantilla antes.")
    if chave_col not in df_cuenta.columns:
        raise ValueError("Cuenta (GL0061) não contém a coluna 'Chave'. Rode a etapa de Cuenta antes.")

    def _norm_key_series(s: pd.Series) -> pd.Series:
        return (
            s.astype(str)
             .str.replace("\u2212", "-", regex=False)
             .str.replace("\xa0", " ", regex=False)
             .str.replace(r"\s+", " ", regex=True)
             .str.strip()
        )

    df_pg = df_pg.copy()
    df_cuenta = df_cuenta.copy()
    df_pg["_key_norm"] = _norm_key_series(df_pg[chave_col])
    df_cuenta["_key_norm"] = _norm_key_series(df_cuenta[chave_col])

    first_original_key = df_pg.groupby("_key_norm")[chave_col].first()

    cnt_pg = df_pg["_key_norm"].value_counts()
    cnt_cuenta = df_cuenta["_key_norm"].value_counts()

    amount_col = None
    for c in df_pg.columns:
        cl = str(c).strip().lower()
        if cl == "amount" or "amount" in cl:
            amount_col = c
            break
    cuenta_amt_col = "Saldo Real" if "Saldo Real" in df_cuenta.columns else None

    sum_pg = pd.Series(dtype=float)
    sum_cuenta = pd.Series(dtype=float)
    if amount_col is not None:
        sum_pg = pd.to_numeric(df_pg[amount_col], errors="coerce").groupby(df_pg["_key_norm"]).sum(min_count=1)
    if cuenta_amt_col is not None:
        sum_cuenta = pd.to_numeric(df_cuenta[cuenta_amt_col], errors="coerce").groupby(df_cuenta["_key_norm"]).sum(min_count=1)

    keep_limit_map = cnt_cuenta.to_dict()

    rank = df_pg.groupby("_key_norm").cumcount() + 1
    keep_limit = df_pg["_key_norm"].map(keep_limit_map)

    mask_keep = keep_limit.isna() | (rank <= keep_limit.fillna(np.inf))

    if not sum_pg.empty and not sum_cuenta.empty:
        keys_ok = []
        for k, cpg in cnt_pg.items():
            cct = int(cnt_cuenta.get(k, 0))
            if cpg == cct and cct > 0:
                spg = float(sum_pg.get(k, 0.0))
                scu = float(sum_cuenta.get(k, 0.0))
                if abs(spg - scu) <= tol_soma:
                    keys_ok.append(k)
        if keys_ok:
            mask_keep = mask_keep | df_pg["_key_norm"].isin(keys_ok)

    df_clean = df_pg[mask_keep].copy().reset_index(drop=True)
    df_clean.drop(columns=["_key_norm"], inplace=True, errors="ignore")

    removed_total = int((~mask_keep).sum())
    removed_by_key = {}
    keys_with_drop = []
    for k, cpg in cnt_pg.items():
        cct = int(cnt_cuenta.get(k, 0))
        if cpg > cct and cct > 0:
            diff = cpg - cct
            k_disp = str(first_original_key.get(k, k))
            removed_by_key[k_disp] = diff
            keys_with_drop.append(k)

    stats = {
        "rows_original": int(len(df_pg)),
        "rows_clean": int(len(df_clean)),
        "rows_removed": removed_total,
        "keys_with_removal": len(keys_with_drop),
        "removed_by_key": dict(sorted(removed_by_key.items(), key=lambda kv: kv[1], reverse=True)),
    }
    return df_clean, stats

# -----------------------------------------------------------------------------
# Parsers - ESTADO DE CUENTA (.txt)
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
    except Exception:
        return None


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
# PARSER GL0061 — fixed-width lines
# -----------------------------------------------------------------------------
def parse_cuenta_gl(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()
    dados = []

    cta_header = None
    reg_header = re.compile(r"Nº de cta\.\s+(\d{6})")
    for ln in linhas[:50]:
        m = reg_header.search(ln)
        if m:
            cta_header = m.group(1)
            break
    if not cta_header:
        raise ValueError("CTA não encontrada no cabeçalho do arquivo GL0061.")

    def clean_num(v: str | None) -> float:
        if not v:
            return 0.0
        s = str(v).strip()
        neg = s.endswith("-")
        if neg:
            s = s[:-1]
        s = s.replace(",", "")
        try:
            val = float(s)
        except Exception:
            val = 0.0
        return -val if neg else val

    ignore = re.compile(
        r"Electrolux|Planificación|Moneda|Scala|^-{3,}|^={3,}|"
        r"Saldo Inicial|Saldo final|T O T A L|ACTIVO|Página|Criterios|CUENTAS POR"
    )

    cols = [
        "CTA","CC","PROD","CNT","TDW",
        "Fecha","Transacción",
        "Debe","Haber",
        "Saldo Real","Saldo",
        "Texto"
    ]

    for ln in linhas:
        if ignore.search(ln):
            continue
        if len(ln.strip()) == 0:
            continue
        if not re.search(r"\d{2}/\d{2}/\d{2}", ln):
            continue

        cc     = ln[0:5].strip()
        prod   = ln[5:13].strip()
        cnt    = ln[13:23].strip()
        tdw    = ln[23:31].strip()
        fecha  = ln[31:40].strip()
        ntran  = ln[40:50].strip()

        nums = re.findall(r"[-\d,]+\.\d{2}-?", ln)
        if len(nums) < 3:
            continue

        debe  = clean_num(nums[-3])
        haber = clean_num(nums[-2])
        saldo_impresso = clean_num(nums[-1])

        saldo_real = round(debe - haber, 2)

        texto_pos = ln.rfind(nums[-1])
        texto = ln[texto_pos + len(nums[-1]):].strip() if texto_pos != -1 else ""

        if not cc.isdigit():
            cc = ""

        dados.append([
            cta_header, cc, prod, cnt, tdw,
            fecha, ntran, debe, haber,
            saldo_real, saldo_impresso,
            texto
        ])

    df = pd.DataFrame(dados, columns=cols)

    for date_col in ["Fecha", "Fechado"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.date

    return df


# -----------------------------------------------------------------------------
# Export XLSX with number/date formats (tuned for big dataframes)
# -----------------------------------------------------------------------------
def to_xlsx_bytes_format(
    df: pd.DataFrame,
    sheet_name: str,
    numeric_cols: list[str] | None = None,
    date_cols: list[str] | None = None,
    *,
    format_cells: bool = True,
    max_autowidth_rows: int = 3000
) -> bytes:
    numeric_cols = numeric_cols or []
    date_cols = date_cols or []

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_to_save = df.copy()
        for dc in date_cols:
            if dc in df_to_save.columns:
                df_to_save[dc] = pd.to_datetime(df_to_save[dc], errors="coerce")

        df_to_save.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        BLUE = "FF0077B6"
        WHITE = "FFFFFFFF"
        fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
        font_white_bold = Font(color=WHITE, bold=True)
        for cell in ws[1]:
            cell.fill = fill_blue
            cell.font = font_white_bold

        if format_cells:
            for col_name in numeric_cols:
                if col_name not in df_to_save.columns:
                    continue
                col_idx = df_to_save.columns.get_loc(col_name) + 1
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    if isinstance(cell.value, (int, float)) and cell.value is not None:
                        cell.number_format = '#,##0.00'

            for col_name in date_cols:
                if col_name not in df_to_save.columns:
                    continue
                col_idx = df_to_save.columns.get_loc(col_name) + 1
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    if cell.value:
                        cell.number_format = 'dd/mm/yyyy'

        # Auto-width limited for speed
        for col_idx in range(1, ws.max_column + 1):
            max_len = 10
            limit = min(ws.max_row, max_autowidth_rows)
            for row in range(1, limit + 1):
                v = ws.cell(row=row, column=col_idx).value
                if v is None:
                    continue
                if isinstance(v, (int, float)):
                    s = f"{v:,.2f}"
                else:
                    s = str(v)
                max_len = max(max_len, len(s))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    buffer.seek(0)
    return buffer.getvalue()
# -----------------------------------------------------------------------------
# Page
# -----------------------------------------------------------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    col_b1, col_b2, col_b3, col_b4, col_b5 = st.columns(5)

    limpeza_pg_clicked = False

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
    with col_b5:
        if st.button("Limpieza Plantilla Gastos", use_container_width=True):
            limpeza_pg_clicked = True

    mode = st.session_state["aag_mode"]
    st.divider()

    # ====== ACTION: Limpieza Plantilla Gastos ======
    if limpeza_pg_clicked:
        st.subheader("🧹 Limpieza da Plantilla de Gastos")
        try:
            df_pg = st.session_state.get("aag_plantilla_df", None)
            df_ct = st.session_state.get("aag_cuenta_df", None)

            if df_pg is None or df_pg.empty:
                st.error("Antes de limpar, carregue e execute a **Plantilla de Gastos**.")
            elif df_ct is None or df_ct.empty:
                st.error("Antes de limpar, carregue e processe o **Archivo de Cuenta (GL0061)**.")
            else:
                df_pg_clean, stats = limpiar_plantilla_contra_cuenta(df_pg, df_ct, chave_col="Chave")
                st.session_state["aag_plantilla_df"] = df_pg_clean.copy()
                st.session_state["aag_state"]["last_action"] = "limpieza_pg"

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Linhas (Original)", f"{stats['rows_original']:,}".replace(",", "."))
                with c2:
                    st.metric("Linhas (Limpo)", f"{stats['rows_clean']:,}".replace(",", "."))
                with c3:
                    st.metric("Removidas", f"{stats['rows_removed']:,}".replace(",", "."))
                with c4:
                    st.metric("Chaves Ajustadas", f"{stats['keys_with_removal']:,}".replace(",", "."))

                if stats["rows_removed"] == 0:
                    st.info("Nenhuma divergência de contagem encontrada. Nada foi removido.")
                else:
                    with st.expander("Ver detalhes por chave (quantidade removida)"):
                        det = pd.DataFrame(
                            [{"Chave": k, "Removidas": v} for k, v in stats["removed_by_key"].items()]
                        ).sort_values(by="Removidas", ascending=False)
                        st.dataframe(det, use_container_width=True, height=280)

                st.caption("Prévia da Plantilla de Gastos (após limpeza):")
                df_pg_prev = st.session_state["aag_plantilla_df"]

                amount_col_view = None
                for c in df_pg_prev.columns:
                    if str(c).strip().lower() == "amount":
                        amount_col_view = c
                        break
                if amount_col_view is None:
                    candidates = [c for c in df_pg_prev.columns if "amount" in str(c).strip().lower()]
                    if candidates:
                        amount_col_view = candidates[0]

                known_date_keys = {"transactiondate","duedate","due_date","invoicedate","invoice_date"}
                found_date_cols_view = [c for c in df_pg_prev.columns if c.lower().replace(" ", "_") in known_date_keys]

                col_cfg = {}
                if amount_col_view:
                    col_cfg[str(amount_col_view)] = st.column_config.NumberColumn(format="%.2f")
                for dc in found_date_cols_view:
                    if pd.api.types.is_datetime64_any_dtype(df_pg_prev[dc]):
                        col_cfg[dc] = st.column_config.DateColumn(format="DD/MM/YYYY")
                    else:
                        col_cfg[dc] = st.column_config.TextColumn()

                st.dataframe(df_pg_prev.head(MAX_PREVIEW_DEFAULT), use_container_width=True, height=520, column_config=col_cfg)

                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        "Baixar CSV (Plantilla Limpia)",
                        df_pg_prev.to_csv(index=False).encode("utf-8"),
                        "plantilla_gastos_limpia.csv",
                        "text/csv",
                        use_container_width=True
                    )
                with col_xlsx:
                    format_cells = len(df_pg_prev) <= XLSX_HARD_LIMIT_ROWS
                    if not format_cells:
                        st.info("XLSX muito grande: gerando sem formatação por célula para evitar travamentos.")
                    xlsx_bytes = to_xlsx_bytes_format(
                        df_pg_prev,
                        sheet_name="PlantillaGastos (Limpia)",
                        numeric_cols=[amount_col_view] if amount_col_view else [],
                        date_cols=[c for c in df_pg_prev.columns if pd.api.types.is_datetime64_any_dtype(df_pg_prev[c]) or c in found_date_cols_view],
                        format_cells=False,
                    )
                    st.download_button(
                        "Baixar XLSX (Plantilla Limpia)",
                        xlsx_bytes,
                        "plantilla_gastos_limpia.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        except Exception as e:
            st.error("Erro durante a limpeza da Plantilla de Gastos.")
            st.exception(e)

    # -------------------------------------------------------------------------
    # Mode: Estado de Cuenta (.txt)
    # -------------------------------------------------------------------------
    if mode == "estado":
        upl_key_estado = st.session_state["aag_state"].setdefault("uploader_key_estado", "aag_estado_upl_1")

        st.caption("Carregue o arquivo **.txt** de *Listado de Saldos* para visualização e export.")
        uploaded = st.file_uploader(
            "Selecionar arquivo (.txt)",
            type=["txt"],
            accept_multiple_files=False,
            key=upl_key_estado,
            help="Ex.: relatório 'Listado de Saldos' exportado do sistema.",
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", use_container_width=True, disabled=(uploaded is None))
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_estado"] = upl_key_estado + "_x"
            if "aag_estado_df" in st.session_state:
                del st.session_state["aag_estado_df"]
            st.rerun()

        if run_clicked and uploaded is not None:
            pbar = st.progress(0, text="Lendo arquivo .txt...")
            try:
                raw_bytes = uploaded.getvalue()
                try:
                    text = raw_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    text = raw_bytes.decode("latin-1")

                pbar.progress(35, text="Convertendo para DataFrame...")
                df_base = parse_estado_cuenta_txt(text)

                if df_base is None or df_base.empty:
                    st.warning("Nenhuma linha válida encontrada no arquivo.")
                    pbar.progress(0, text="Aguardando...")
                    return

                st.session_state["aag_estado_df"] = df_base.copy()

                pbar.progress(70, text="Preparando visualização...")
                st.success("Arquivo processado com sucesso.")
                pbar.progress(100, text="Concluído.")
            except Exception as e:
                st.error("Erro ao processar o arquivo .txt.")
                st.exception(e)

        if "aag_estado_df" in st.session_state and isinstance(st.session_state["aag_estado_df"], pd.DataFrame):
            df_base = st.session_state["aag_estado_df"]

            df = df_base.copy()
            numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
            for c in numeric_cols:
                df[c] = pd.to_numeric(df[c], errors="coerce")
            totals = {c: float(np.nansum(df[c].values)) for c in numeric_cols}
            total_row = {col: "" for col in df.columns}
            total_row["Descripción"] = "TOTAL"
            for c in numeric_cols:
                total_row[c] = totals[c]
            df = pd.concat([df, pd.DataFrame([total_row], columns=df.columns)], ignore_index=True)

            st.dataframe(
                df,
                use_container_width=True,
                height=550,
                column_config={c: st.column_config.NumberColumn(format="%.2f") for c in numeric_cols if c in df.columns},
            )

            col_csv, col_xlsx = st.columns(2)
            with col_csv:
                st.download_button(
                    label="Baixar CSV (Estado de Cuenta)",
                    data=df.to_csv(index=False).encode("utf-8"),
                    file_name="estado_de_cuenta.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with col_xlsx:
                xlsx_bytes = to_xlsx_bytes_format(
                    df, sheet_name="EstadoCuenta", numeric_cols=numeric_cols, date_cols=[]
                )
                st.download_button(
                    label="Baixar XLSX (Estado de Cuenta)",
                    data=xlsx_bytes,
                    file_name="estado_de_cuenta.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    # -------------------------------------------------------------------------
    # Mode: Plantilla de Gastos (.xlsx/.xls) — LOW-MEMORY
    # -------------------------------------------------------------------------
    elif mode == "plantilla":
        upl_key_pg = st.session_state["aag_state"].setdefault("uploader_key_pg", "aag_pg_upl_1")

        st.caption("Carregue o arquivo **Excel** da *Plantilla de Gastos* (primeira aba será lida).")
        uploaded_xl = st.file_uploader(
            "Selecionar arquivo (.xlsx ou .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key=upl_key_pg,
            help="O app guarda apenas colunas essenciais (Cuenta, TransactionDate, TransactionNo, Amount e Chave) para não estourar a memória.",
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", use_container_width=True, disabled=(uploaded_xl is None))
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_pg"] = upl_key_pg + "_x"
            for k in ("aag_plantilla_df", "aag_plantilla_bytes", "aag_plantilla_cols"):
                if k in st.session_state:
                    del st.session_state[k]
            gc.collect()
            st.rerun()

        if run_clicked and uploaded_xl is not None:
            pbar = st.progress(0, text="Lendo arquivo Excel...")
            try:
                name = getattr(uploaded_xl, "name", "").lower()
                engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
                raw_bytes = uploaded_xl.getvalue()

                # Read full sheet once; do not keep full DF in session
                df_full = pd.read_excel(BytesIO(raw_bytes), sheet_name=0, engine=engine)
                df_full.columns = df_full.columns.astype(str).str.strip()
                pbar.progress(30, text="Detectando colunas...")

                def norm(s: str) -> str:
                    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

                amount_col = None
                for c in df_full.columns:
                    if norm(c) == "amount" or "amount" in norm(c):
                        amount_col = c
                        break
                if amount_col is None:
                    st.error("Coluna 'Amount' não encontrada no arquivo.")
                    return

                # Date targets
                date_targets = {"transactiondate": None, "duedate": None, "invoicedate": None}
                for c in df_full.columns:
                    nc = norm(c)
                    if nc == "transactiondate":
                        date_targets["transactiondate"] = c
                    elif nc in ("duedate", "due_date"):
                        date_targets["duedate"] = c
                    elif nc in ("invoicedate", "invoice_date"):
                        date_targets["invoicedate"] = c
                tdate_col = date_targets.get("transactiondate")

                # TransactionNo
                tno_col = None
                for c in df_full.columns:
                    nc = norm(c)
                    if "transaction" in nc and ("no" in nc or "number" in nc):
                        tno_col = c
                        break

                # Cuenta
                cuenta_col = None
                for c in df_full.columns:
                    if norm(c) == "cuenta":
                        cuenta_col = c
                        break
                if cuenta_col is None:
                    st.error("Coluna 'Cuenta' não encontrada.")
                    return

                pbar.progress(55, text="Normalizando tipos (vetorizado)...")

                amount_num = pd.to_numeric(df_full[amount_col], errors="coerce").fillna(0)

                if tdate_col:
                    tdate_dt = pd.to_datetime(df_full[tdate_col], errors="coerce")
                    tdate_str = tdate_dt.dt.strftime("%d/%m/%Y").fillna("")
                else:
                    tdate_dt = pd.Series(pd.NaT, index=df_full.index)
                    tdate_str = pd.Series([""] * len(df_full), index=df_full.index)

                if tno_col:
                    tno_series = (
                        df_full[tno_col]
                        .astype(str)
                        .str.replace(r"\D", "", regex=True)
                    )
                    tno_series = tno_series.where(tno_series.str.len() > 0, "")
                    tno_series = tno_series.mask(tno_series == "", None)
                    tno_series = pd.to_numeric(tno_series, errors="coerce").astype("Int64").astype(str)
                    tno_series = tno_series.replace("<NA>", "")
                    tno_series = tno_series.str.zfill(9)
                else:
                    tno_series = pd.Series([""] * len(df_full), index=df_full.index)

                cuenta_str = df_full[cuenta_col].astype(str).str.strip().fillna("")

                pbar.progress(75, text="Construindo Chave (núcleo leve)...")
                chave = cuenta_str + "|" + tdate_str + "|" + tno_series + "|" + amount_num.map("{:.2f}".format)

                df_core = pd.DataFrame({
                    cuenta_col: df_full[cuenta_col],
                    (tdate_col if tdate_col else "TransactionDate"): tdate_dt,
                    (tno_col if tno_col else "TransactionNo"): tno_series,
                    amount_col: amount_num,
                    "Chave": chave,
                })

                st.session_state["aag_plantilla_df"] = df_core
                st.session_state["aag_plantilla_bytes"] = raw_bytes
                st.session_state["aag_plantilla_cols"] = {
                    "cuenta": cuenta_col,
                    "tdate": tdate_col if tdate_col else "TransactionDate",
                    "tno": tno_col if tno_col else "TransactionNo",
                    "amount": amount_col,
                }

                del df_full
                gc.collect()

                st.success("Arquivo carregado com sucesso (núcleo em memória).")
                pbar.progress(100, text="Concluído.")
            except Exception as e:
                st.error("Erro ao processar o arquivo Excel.")
                st.exception(e)

        # Display & downloads based on core DF only
        if "aag_plantilla_df" in st.session_state and isinstance(st.session_state["aag_plantilla_df"], pd.DataFrame):
            df_pg = st.session_state["aag_plantilla_df"]
            cols_map = st.session_state.get("aag_plantilla_cols", {})
            cuenta_col = cols_map.get("cuenta")
            tdate_col  = cols_map.get("tdate")
            tno_col    = cols_map.get("tno")
            amount_col = cols_map.get("amount")

            total_rows = len(df_pg)
            if total_rows > 2000:
                preview_rows = st.slider(
                    "Linhas na pré-visualização (núcleo)",
                    min_value=1000, max_value=min(total_rows, MAX_PREVIEW_MAX),
                    value=min(total_rows, MAX_PREVIEW_DEFAULT), step=1000
                )
            else:
                preview_rows = total_rows

            if total_rows > preview_rows:
                st.caption(f"Mostrando as primeiras {preview_rows:,} de {total_rows:,} linhas (apenas colunas essenciais).")

            df_preview = df_pg.iloc[:preview_rows]

            col_cfg = {}
            if amount_col and amount_col in df_preview.columns:
                col_cfg[str(amount_col)] = st.column_config.NumberColumn(format="%.2f")
            if tdate_col and tdate_col in df_preview.columns and pd.api.types.is_datetime64_any_dtype(df_preview[tdate_col]):
                col_cfg[tdate_col] = st.column_config.DateColumn(format="DD/MM/YYYY")

            st.dataframe(df_preview, use_container_width=True, height=550, column_config=col_cfg)

            col_csv_core, col_xlsx_core = st.columns(2)

            with col_csv_core:
                st.download_button(
                    label="Baixar CSV (núcleo leve)",
                    data=df_pg.to_csv(index=False).encode("utf-8"),
                    file_name="plantilla_gastos_nucleo.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            with col_xlsx_core:
                if total_rows > XLSX_HARD_LIMIT_ROWS:
                    st.info(f"Base com {total_rows:,} linhas — XLSX desabilitado para evitar estouro de memória. Use o CSV.")
                else:
                    xlsx_bytes = to_xlsx_bytes_format(
                        df_pg,
                        sheet_name="PlantillaGastos (Núcleo)",
                        numeric_cols=[amount_col] if amount_col else [],
                        date_cols=[tdate_col] if tdate_col else [],
                        format_cells=False,
                    )
                    st.download_button(
                        label="Baixar XLSX (núcleo leve)",
                        data=xlsx_bytes,
                        file_name="plantilla_gastos_nucleo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            st.divider()
            with st.expander("⚠️ Opções avançadas (podem usar muita memória)"):
                enable_full = st.toggle("Habilitar download COMPLETO (lento e pode travar)", value=False)
                if enable_full:
                    if st.button("Gerar download COMPLETO (CSV)", type="secondary"):
                        try:
                            raw = st.session_state.get("aag_plantilla_bytes", None)
                            if raw is None:
                                st.warning("Bytes originais não encontrados na sessão.")
                            else:
                                engine = "openpyxl"  # assuming .xlsx (adjust if needed)
                                df_full = pd.read_excel(BytesIO(raw), sheet_name=0, engine=engine)
                                csv_bytes = df_full.to_csv(index=False).encode("utf-8")
                                st.download_button(
                                    "Baixar CSV COMPLETO",
                                    data=csv_bytes,
                                    file_name="plantilla_completa.csv",
                                    mime="text/csv",
                                    use_container_width=True,
                                )
                                del df_full, csv_bytes
                                gc.collect()
                        except Exception as e:
                            st.error("Falha ao gerar o CSV completo.")
                            st.exception(e)

    # -------------------------------------------------------------------------
    # Mode: Analise — compare Estado de Cuenta (CTA/Período) x Plantilla (Cuenta/Amount)
    # -------------------------------------------------------------------------
    elif mode == "asientos":
        st.subheader("🔍 Analise: Estado de Cuenta x Plantilla de Gastos")

        df_ec = st.session_state.get("aag_estado_df", None)
        df_pg = st.session_state.get("aag_plantilla_df", None)

        missing = []
        if df_ec is None or df_ec.empty:
            missing.append("Estado de Cuenta (.txt)")
        if df_pg is None or df_pg.empty:
            missing.append("Plantilla de Gastos (.xlsx/.xls)")

        if missing:
            st.warning(
                "Para executar a análise, primeiro carregue: " + ", ".join(missing) +
                ". Use as abas **Estado de Cuenta** e **Plantilla Gastos**."
            )
            return

        tol = st.number_input("Valor de Tolerância", min_value=0.00, value=0.01, step=0.01)

        def _norm_conta(x) -> str:
            s = re.sub(r"\D", "", str(x))
            s = s.lstrip("0")
            return s if s else ""

        pbar = st.progress(0, text="Consolidando Estado de Cuenta...")

        try:
            if "CTA" not in df_ec.columns or "Período" not in df_ec.columns:
                st.error("Estado de Cuenta não contém as colunas esperadas: 'CTA' e 'Período'.")
                return

            df_ec_proc = df_ec.copy()
            df_ec_proc["CTA"] = df_ec_proc["CTA"].apply(_norm_conta)
            df_ec_proc["Período"] = pd.to_numeric(df_ec_proc["Período"], errors="coerce").fillna(0.0)
            df_ec_proc = df_ec_proc[df_ec_proc["CTA"].astype(str).str.len() > 0]

            df_ec_agg = (
                df_ec_proc.groupby("CTA", as_index=False)["Período"]
                .sum()
                .rename(columns={"CTA": "Cuenta", "Período": "Saldo_Estado_Cuenta"})
            )

            pbar.progress(40, text="Consolidando Plantilla de Gastos...")
        except Exception as e:
            st.error("Erro ao consolidar Estado de Cuenta.")
            st.exception(e)
            return

        try:
            def _find_col(df, target):
                for c in df.columns:
                    if str(c).strip().lower() == target:
                        return c
                cand = [c for c in df.columns if target in str(c).strip().lower()]
                return cand[0] if cand else None

            cuenta_col = _find_col(df_pg, "cuenta")
            amount_col = _find_col(df_pg, "amount")

            if cuenta_col is None or amount_col is None:
                st.error("Plantilla não contém as colunas esperadas: 'Cuenta' e 'Amount'.")
                return

            df_pg_proc = df_pg.copy()
            df_pg_proc[amount_col] = pd.to_numeric(df_pg_proc[amount_col], errors="coerce").fillna(0.0)
            df_pg_proc["__conta__"] = df_pg_proc[cuenta_col].apply(_norm_conta)
            df_pg_proc = df_pg_proc[df_pg_proc["__conta__"].astype(str).str.len() > 0]

            df_pg_agg = (
                df_pg_proc.groupby("__conta__", as_index=False)[amount_col]
                .sum()
                .rename(columns={"__conta__": "Cuenta", amount_col: "Saldo_Plantilla_Gastos"})
            )

            pbar.progress(70, text="Comparando saldos...")
        except Exception as e:
            st.error("Erro ao consolidar Plantilla de Gastos.")
            st.exception(e)
            return

        try:
            df_cmp = pd.merge(df_ec_agg, df_pg_agg, on="Cuenta", how="outer")
            for c in ["Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos"]:
                df_cmp[c] = pd.to_numeric(df_cmp[c], errors="coerce").fillna(0.0)

            df_cmp["Diferença"] = (df_cmp["Saldo_Plantilla_Gastos"] - df_cmp["Saldo_Estado_Cuenta"]).round(2)

            df_cmp["_div"] = df_cmp["Diferença"].abs() > float(tol)

            df_cmp["_cuenta_num"] = pd.to_numeric(df_cmp["Cuenta"], errors="coerce")
            df_cmp = df_cmp.sort_values(by="_cuenta_num", ascending=True).drop(columns=["_cuenta_num"]).reset_index(drop=True)

            pbar.progress(90, text="Preparando visualização...")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Contas (Estado)", f"{df_ec_agg['Cuenta'].nunique():,}".replace(",", "."))
            with c2:
                st.metric(
                    "Soma Estado de Cuentas",
                    f"{df_ec_agg['Saldo_Estado_Cuenta'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                )
            with c3:
                st.metric(
                    "Soma Plantilla de Gastos",
                    f"{df_pg_agg['Saldo_Plantilla_Gastos'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                )

            only_div = st.checkbox("Mostrar apenas contas com divergência", value=True)
            df_show = df_cmp[df_cmp["_div"]] if only_div else df_cmp

            df_show = df_show[["Cuenta", "Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos", "Diferença"]]

            st.dataframe(
                df_show,
                use_container_width=True, height=520,
                column_config={
                    "Saldo_Estado_Cuenta": st.column_config.NumberColumn(format="%.2f"),
                    "Saldo_Plantilla_Gastos": st.column_config.NumberColumn(format="%.2f"),
                    "Diferença": st.column_config.NumberColumn(format="%.2f"),
                },
            )

            pbar.progress(95, text="Gerando arquivos para download...")

            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label="Baixar CSV (Analise)",
                    data=df_show.to_csv(index=False).encode("utf-8"),
                    file_name="analise_contas.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with col_d2:
                xlsx_bytes = to_xlsx_bytes_format(
                    df_show,
                    sheet_name="Analise",
                    numeric_cols=["Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos", "Diferença"],
                    date_cols=[],
                )
                st.download_button(
                    label="Baixar XLSX (Analise)",
                    data=xlsx_bytes,
                    file_name="analise_contas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            pbar.progress(100, text="Concluído.")

        except Exception as e:
            st.error("Erro durante a comparação.")
            st.exception(e)

    # -------------------------------------------------------------------------
    # Mode: Cuenta (GL0061)
    # -------------------------------------------------------------------------
    elif mode == "cuenta":
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
            except Exception:
                text = raw.decode("latin-1")

            try:
                df = parse_cuenta_gl(text)
            except Exception as e:
                st.error("Erro ao interpretar o arquivo GL0061.")
                st.exception(e)
                return

            if df.empty:
                st.error("Nenhuma linha reconhecida no arquivo GL0061.")
                return

            cta_str   = df["CTA"].apply(_str_or_empty) if "CTA" in df.columns else pd.Series([""]*len(df))
            fecha_str = df["Fecha"].apply(_fmt_date_ddmmyyyy) if "Fecha" in df.columns else pd.Series([""]*len(df))
            tran_str  = df["Transacción"].apply(_fmt_transno_keep_zeros) if "Transacción" in df.columns else pd.Series([""]*len(df))
            sreal_str = df["Saldo Real"].apply(_fmt_num_2dec_point) if "Saldo Real" in df.columns else pd.Series([""]*len(df))

            df["Chave"] = cta_str + "|" + fecha_str + "|" + tran_str + "|" + sreal_str

            st.session_state["aag_cuenta_df"] = df.copy()

        if "aag_cuenta_df" in st.session_state and isinstance(st.session_state["aag_cuenta_df"], pd.DataFrame):
            df = st.session_state["aag_cuenta_df"]

            date_cols = [c for c in ["Fecha", "Fechado"] if c in df.columns]
            col_cfg = {
                "Debe": st.column_config.NumberColumn(format="%.2f"),
                "Haber": st.column_config.NumberColumn(format="%.2f"),
                "Saldo Real": st.column_config.NumberColumn(format="%.2f"),
                "Saldo": st.column_config.NumberColumn(format="%.2f"),
            }
            for dc in date_cols:
                col_cfg[dc] = st.column_config.DateColumn(format="DD/MM/YYYY")

            st.dataframe(df, use_container_width=True, height=600, column_config=col_cfg)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "Baixar CSV",
                    df.to_csv(index=False).encode("utf-8"),
                    "cuenta.csv",
                    "text/csv",
                    use_container_width=True
                )
            with col2:
                xlsx_bytes = to_xlsx_bytes_format(
                    df, "Cuenta",
                    numeric_cols=["Debe","Haber","Saldo Real","Saldo"],
                    date_cols=date_cols
                )
                st.download_button(
                    "Baixar XLSX",
                    xlsx_bytes,
                    "cuenta.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    else:
        st.info("Selecione um modo acima para continuar.")

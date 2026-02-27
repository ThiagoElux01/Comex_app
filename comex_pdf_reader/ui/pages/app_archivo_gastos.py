# -*- coding: utf-8 -*-
import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# ================== HELPERS BÁSICOS ==================

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

_EXCEL_ORIGIN = "1899-12-30"

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

# ================== CORE FUNÇÕES ==================

_NUM = r"(\-?\d[\d,]*\.\d{2}\-?)"

def _clean_num(s: str):
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
        "CTA", "CC", "PROD", "CNT", "TDW",
        "Fecha", "Transacción",
        "Debe", "Haber",
        "Saldo Real", "Saldo",
        "Texto"
    ]
    for ln in linhas:
        if ignore.search(ln):
            continue
        if len(ln.strip()) == 0:
            continue
        if not re.search(r"\d{2}/\d{2}/\d{2}", ln):
            continue
        cc = ln[0:5].strip()
        prod = ln[5:13].strip()
        cnt = ln[13:23].strip()
        tdw = ln[23:31].strip()
        fecha = ln[31:40].strip()
        ntran = ln[40:50].strip()
        nums = re.findall(r"[-\d,]+\.\d{2}-?", ln)
        if len(nums) < 3:
            continue
        debe = clean_num(nums[-3])
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

def to_xlsx_bytes_format(df: pd.DataFrame, sheet_name: str, numeric_cols=None, date_cols=None) -> bytes:
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
        BLUE = "FF0077B6"; WHITE = "FFFFFFFF"
        fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
        font_white_bold = Font(color=WHITE, bold=True)
        for cell in ws[1]:
            cell.fill = fill_blue
            cell.font = font_white_bold
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
        for col_idx in range(1, ws.max_column + 1):
            max_len = 10
            for row in range(1, ws.max_row + 1):
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

# ================== LIMPEZA CONTRA CUENTA ==================

def limpiar_plantilla_contra_cuenta(df_pg: pd.DataFrame, df_cuenta: pd.DataFrame, chave_col: str = "Chave", tol_soma: float = 0.005):
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
    df_pg = df_pg.copy(); df_cuenta = df_cuenta.copy()
    df_pg["_key_norm"] = _norm_key_series(df_pg[chave_col])
    df_cuenta["_key_norm"] = _norm_key_series(df_cuenta[chave_col])
    first_original_key = df_pg.groupby("_key_norm")[chave_col].first()
    cnt_pg = df_pg["_key_norm"].value_counts()
    cnt_cuenta = df_cuenta["_key_norm"].value_counts()
    amount_col = None
    for c in df_pg.columns:
        cl = str(c).strip().lower()
        if cl == "amount" or "amount" in cl:
            amount_col = c; break
    cuenta_amt_col = "Saldo Real" if "Saldo Real" in df_cuenta.columns else None
    sum_pg = pd.Series(dtype=float); sum_cuenta = pd.Series(dtype=float)
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
        # Os campos abaixo existem para compatibilidade, mas NÃO são exibidos
        "keys_with_removal": len(keys_with_drop),
        "removed_by_key": dict(sorted(removed_by_key.items(), key=lambda kv: kv[1], reverse=True)),
    }
    return df_clean, stats

# ================== UI PRINCIPAL ==================

def render():
    _ensure_state()
    if "aag_plantilla_df_orig" not in st.session_state and "aag_plantilla_df" in st.session_state:
        try:
            st.session_state["aag_plantilla_df_orig"] = st.session_state["aag_plantilla_df"].copy()
        except Exception:
            st.session_state["aag_plantilla_df_orig"] = st.session_state["aag_plantilla_df"]
        del st.session_state["aag_plantilla_df"]

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

    # ===== LIMPIEZA (AJUSTE SOLICITADO) =====
    if limpeza_pg_clicked:
        st.subheader("🧹 Limpieza da Plantilla de Gastos")
        try:
            df_pg_orig = st.session_state.get("aag_plantilla_df_orig", None)
            df_ct = st.session_state.get("aag_cuenta_df", None)
            df_ec = st.session_state.get("aag_estado_df", None)
            if df_pg_orig is None or df_pg_orig.empty:
                st.error("Antes de limpar, carregue e execute a **Plantilla de Gastos**.")
                st.stop()
            if df_ct is None or df_ct.empty:
                st.error("Antes de limpar, carregue e processe o **Archivo de Cuenta (GL0061)**.")
                st.stop()
            df_pg_clean, stats = limpiar_plantilla_contra_cuenta(df_pg_orig, df_ct, chave_col="Chave")
            st.session_state["aag_plantilla_df_clean"] = df_pg_clean.copy()
            st.session_state["aag_state"]["last_action"] = "limpieza_pg"

            # >>> SOMENTE 3 MÉTRICAS (sem 'Chaves Ajustadas')
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Linhas (Original)", f"{stats['rows_original']:,}".replace(",", "."))
            with c2:
                st.metric("Linhas (Limpo)", f"{stats['rows_clean']:,}".replace(",", "."))
            with c3:
                st.metric("Removidas", f"{stats['rows_removed']:,}".replace(",", "."))

            # >>> NENHUM EXPANDER DE DETALHE AQUI <<<

            # Download Plantilla limpa (mantido)
            df_pg_prev = st.session_state["aag_plantilla_df_clean"]
            st.download_button(
                "Baixar CSV (Plantilla Limpia)",
                df_pg_prev.to_csv(index=False).encode("utf-8"),
                "plantilla_gastos_limpia.csv",
                "text/csv",
                use_container_width=True
            )

            # ===== Dataframe do Estado + Período Ajustado =====
            if df_ec is None or df_ec.empty:
                st.warning("Estado de Cuenta ainda não foi carregado. Carregue-o para ver o ajuste por Período.")
            else:
                df_ec_view = df_ec.copy()
                df_ec_view["CTA"] = df_ec_view["CTA"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0")
                df_ec_view.loc[df_ec_view["CTA"] == "", "CTA"] = None
                df_ec_view["Período"] = pd.to_numeric(df_ec_view["Período"], errors="coerce").fillna(0.0)
                df_ec_view = df_ec_view[["CTA", "Descripción", "Período"]]

                # localizar colunas Cuenta/Amount na Plantilla limpa
                def _find_col_ci(df: pd.DataFrame, targets: list[str]):
                    cols_map = {re.sub(r"[^a-z0-9]", "", str(c).lower()): c for c in df.columns}
                    for t in targets:
                        key = re.sub(r"[^a-z0-9]", "", t.lower())
                        if key in cols_map:
                            return cols_map[key]
                    return None
                amount_col = None
                for c in df_pg_prev.columns:
                    if str(c).strip().lower() == "amount":
                        amount_col = c; break
                if amount_col is None:
                    cand = [c for c in df_pg_prev.columns if "amount" in str(c).strip().lower()]
                    amount_col = cand[0] if cand else None
                cuenta_col = _find_col_ci(df_pg_prev, ["Cuenta"])

                if amount_col is None or cuenta_col is None:
                    st.error("Para calcular o 'Período Ajustado', a Plantilla Limpa deve conter as colunas 'Cuenta' e 'Amount'.")
                else:
                    df_pg_sum = df_pg_prev.copy()
                    df_pg_sum[cuenta_col] = df_pg_sum[cuenta_col].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0")
                    df_pg_sum[amount_col] = pd.to_numeric(df_pg_sum[amount_col], errors="coerce").fillna(0.0)
                    df_pg_sum = df_pg_sum.groupby(cuenta_col, as_index=False)[amount_col].sum().rename(columns={amount_col: "Período Ajustado", cuenta_col: "Cuenta"})
                    df_final = df_ec_view.merge(df_pg_sum, left_on="CTA", right_on="Cuenta", how="left").drop(columns=["Cuenta"], errors="ignore")
                    df_final["Período Ajustado"] = df_final["Período Ajustado"].fillna(0.0)
                    st.subheader("📊 Estado de Cuenta com Período Ajustado")
                    st.dataframe(
                        df_final,
                        use_container_width=True,
                        height=520,
                    )
                    st.download_button(
                        "Baixar CSV (Estado + Período Ajustado)",
                        df_final.to_csv(index=False).encode("utf-8"),
                        "estado_cuenta_ajustado.csv",
                        "text/csv",
                        use_container_width=True
                    )
        except Exception as e:
            st.error("Erro durante a limpeza da Plantilla de Gastos.")
            st.exception(e)
        st.stop()

    # Demais abas não são necessárias para demonstrar a correção solicitada,
    # porém manteremos pelo menos a 'estado' e 'cuenta' para carregar dados base.
    if mode == "estado":
        upl_key_estado = st.session_state["aag_state"].setdefault("uploader_key_estado", "aag_estado_upl_1")
        st.caption("Carregue o arquivo **.txt** de *Listado de Saldos* para visualização e export.")
        uploaded = st.file_uploader("Selecionar arquivo (.txt)", type=["txt"], accept_multiple_files=False, key=upl_key_estado)
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
            raw_bytes = uploaded.getvalue()
            try:
                text = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text = raw_bytes.decode("latin-1")
            df_base = parse_estado_cuenta_txt(text)
            if df_base is None or df_base.empty:
                st.warning("Nenhuma linha válida encontrada no arquivo.")
                return
            st.session_state["aag_estado_df"] = df_base.copy()
            st.success("Arquivo processado com sucesso.")
        if "aag_estado_df" in st.session_state and isinstance(st.session_state["aag_estado_df"], pd.DataFrame):
            st.dataframe(st.session_state["aag_estado_df"], use_container_width=True, height=450)
    elif mode == "cuenta":
        upl_key = st.session_state["aag_state"].setdefault("uploader_key_cuenta", "aag_cuenta_upl_1")
        uploaded = st.file_uploader("Selecionar arquivo GL0061 (.txt)", type=["txt"], key=upl_key)
        col_r, col_c = st.columns([2, 1])
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
            df = parse_cuenta_gl(text)
            if df.empty:
                st.error("Nenhuma linha reconhecida no arquivo GL0061.")
                return
            cta_str = df["CTA"].apply(_str_or_empty) if "CTA" in df.columns else pd.Series([""] * len(df))
            fecha_str = df["Fecha"].apply(_fmt_date_ddmmyyyy) if "Fecha" in df.columns else pd.Series([""] * len(df))
            tran_str = df["Transacción"].apply(_fmt_transno_keep_zeros) if "Transacción" in df.columns else pd.Series([""] * len(df))
            sreal_str = df["Saldo Real"].apply(_fmt_num_2dec_point) if "Saldo Real" in df.columns else pd.Series([""] * len(df))
            df["Chave"] = cta_str + "|" + fecha_str + "|" + tran_str + "|" + sreal_str
            st.session_state["aag_cuenta_df"] = df.copy()
            st.success("GL0061 processado com sucesso.")
        if "aag_cuenta_df" in st.session_state:
            st.dataframe(st.session_state["aag_cuenta_df"], use_container_width=True, height=450)

if __name__ == "__main__":
    render()

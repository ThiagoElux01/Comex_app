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
    """
    Garante que todas as chaves necessárias existam em st.session_state,
    mesmo se já houver um dicionário 'aag_state' antigo/incompleto na sessão.
    """
    # Cria o dicionário de estado principal se não existir
    if "aag_state" not in st.session_state or not isinstance(st.session_state["aag_state"], dict):
        st.session_state["aag_state"] = {}
    aag = st.session_state["aag_state"]

    # Keys dos uploaders separadas por modo (evita conflito de cache do Streamlit)
    aag.setdefault("uploader_key_estado", "aag_estado_upl_1")
    aag.setdefault("uploader_key_pg", "aag_pg_upl_1")
    aag.setdefault("uploader_key_cuenta", "aag_cuenta_upl_1")

    # Última ação (reserva)
    aag.setdefault("last_action", None)

    # Modo atual da página
    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"  # default

def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode

# ==== Helpers de formatação para 'Chave' ====
def _fmt_date_ddmmyyyy(value) -> str:
    """Converte vários tipos de data para 'dd/mm/aaaa' como string."""
    if pd.isna(value):
        return ""
    try:
        # já pode vir como datetime.date, datetime64, string, etc.
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return ""

def _fmt_num_2dec_point(value) -> str:
    """Formata número com 2 casas, ponto como decimal, sem milhares (ex.: 1234.50)."""
    try:
        f = float(value)
        return f"{f:.2f}"
    except Exception:
        return ""

def _str_or_empty(x) -> str:
    return "" if x is None or (isinstance(x, float) and np.isnan(x)) else str(x).strip()

# -----------------------------------------------------------------------------
# Parsers - ESTADO DE CUENTA (.txt)
# -----------------------------------------------------------------------------
_NUM = r"(\-?\d[\d,]*\.\d{2}\-?)"  # número com milhares e 2 decimais; pode terminar com '-' (negativo)

def _clean_num(s: str) -> float | None:
    """Converte strings como '12,345.67-' em float (negativo)."""
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
    """
    Lê um relatório 'Listado de Saldos' em texto e retorna um DataFrame com:
    ['CTA','Descripción','Sal OB','Saldo OB','Período','Saldo CB']
    """
    linhas = texto.splitlines()

    # Encontrar início após o cabeçalho (linha que contém "CTA Descripción")
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
        # Ignora separadores e cabeçalho/rodapé
        if set(raw.strip()) in [{"="}, {"-"}] or "Scala" in raw or "Electrolux" in raw:
            continue

        m = tail_re.search(raw)
        if not m:
            continue

        # Parte à esquerda dos 4 números
        left = raw[: m.start()].rstrip()
        if not left:
            continue

        # CTA = primeiro token; Descripción = resto
        parts = left.split()
        cta = parts[0] if parts else ""
        descr = left[len(cta):].strip() if parts else left.strip()

        # Extrai e normaliza números
        sal_ob, saldo_ob, periodo, saldo_cb = (_clean_num(x) for x in m.groups())
        dados.append([cta, descr, sal_ob, saldo_ob, periodo, saldo_cb])

    cols = ["CTA", "Descripción", "Sal OB", "Saldo OB", "Período", "Saldo CB"]
    df = pd.DataFrame(dados, columns=cols)

    # Tipos numéricos garantidos
    for c in ["Sal OB", "Saldo OB", "Período", "Saldo CB"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# -----------------------------------------------------------------------------
# PARSER GL0061 — colunas fixas
# -----------------------------------------------------------------------------
def parse_cuenta_gl(texto: str) -> pd.DataFrame:
    """
    Parser para arquivos GL0061 (colunas fixas em linha), com:
    CTA | CC | PROD | CNT | TDW | Fecha | Transacción | Debe | Haber | Saldo Real | Saldo | Texto
    """
    linhas = texto.splitlines()
    dados = []

    # Encontrar CTA no cabeçalho (ex.: "Nº de cta. 123456")
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
        return float(v.replace(",", ""))

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

        # Offsets conforme layout fixo do GL0061 (ajuste se necessário)
        cc     = ln[0:5].strip()
        prod   = ln[5:13].strip()
        cnt    = ln[13:23].strip()
        tdw    = ln[23:31].strip()
        fecha  = ln[31:40].strip()
        ntran  = ln[40:50].strip()

        # Últimos 3 números = Debe, Haber, Saldo impresso
        nums = re.findall(r"[-\d,]+\.\d{2}", ln)
        if len(nums) < 3:
            continue

        debe  = clean_num(nums[-3])
        haber = clean_num(nums[-2])
        saldo_impresso = clean_num(nums[-1])

        # Saldo Real = Debe - Haber
        saldo_real = round(debe - haber, 2)

        # Texto após o saldo impresso
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

    # Ajusta datas
    for date_col in ["Fecha", "Fechado"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.date

    return df

# -----------------------------------------------------------------------------
# Export XLSX com máscara numérica e data
# -----------------------------------------------------------------------------
def to_xlsx_bytes_format(
    df: pd.DataFrame,
    sheet_name: str,
    numeric_cols: list[str] | None = None,
    date_cols: list[str] | None = None
) -> bytes:
    """
    Exporta para XLSX aplicando:
      - número: #,##0.00
      - data: dd/mm/yyyy
    """
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

# -----------------------------------------------------------------------------
# ==== NOVOS HELPERS p/ atualização da Plantilla com base em Cuenta ====
# -----------------------------------------------------------------------------
def _find_col_ci(df: pd.DataFrame, targets: list[str]):
    """Busca coluna ignorando acentos/caixa/caracteres especiais."""
    if df is None or df.empty:
        return None
    cols_map = {re.sub(r"[^a-z0-9]", "", str(c).lower()): c for c in df.columns}
    for t in targets:
        key = re.sub(r"[^a-z0-9]", "", t.lower())
        if key in cols_map:
            return cols_map[key]
    return None

def build_plantilla_atualizada_com_cuenta(
    df_pg: pd.DataFrame,
    df_cu: pd.DataFrame,
    tol: float = 0.01
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna (df_pg_atualizada, df_resumo).
    - Ajusta somente as chaves da CTA carregada (chaves que começam com 'CTA|').
    - Para chaves sem diferença de saldo (|sum_pg - sum_cu| <= tol): mantém como está.
    - Para chaves com diferença: recria o bloco na Plantilla espelhando a contagem e valores do GL (Cuenta).
    """
    if df_pg is None or df_pg.empty:
        raise ValueError("Plantilla de Gastos não carregada.")
    if df_cu is None or df_cu.empty:
        raise ValueError("Cuenta (GL0061) não carregada.")

    if "Chave" not in df_pg.columns or "Chave" not in df_cu.columns:
        raise ValueError("Ambos os dataframes precisam ter a coluna 'Chave'.")

    # Detecta colunas essenciais na Plantilla
    amount_col = _find_col_ci(df_pg, ["Amount"])
    cuenta_col = _find_col_ci(df_pg, ["Cuenta"])
    tdate_col  = _find_col_ci(df_pg, ["TransactionDate", "Transaction Date", "TransDate"])
    tno_col    = _find_col_ci(df_pg, ["TransactionNo", "Transaction No", "TransNo", "Transaction_Number"])

    if amount_col is None:
        raise ValueError("Coluna 'Amount' não encontrada na Plantilla.")

    # CTA do GL0061 (é único no arquivo)
    cta_loaded = str(df_cu["CTA"].iloc[0]).strip()
    cta_prefix = f"{cta_loaded}|"

    # Subconjuntos restritos à CTA
    df_pg_subset = df_pg[df_pg["Chave"].astype(str).str.startswith(cta_prefix)].copy()
    df_pg_other  = df_pg[~df_pg["Chave"].astype(str).str.startswith(cta_prefix)].copy()
    df_cu_subset = df_cu[df_cu["Chave"].astype(str).str.startswith(cta_prefix)].copy()

    # Se não houver linhas na Plantilla para esta CTA, usa molde genérico
    template_base = (df_pg_subset.iloc[0].copy() if not df_pg_subset.empty else pd.Series({c: None for c in df_pg.columns}))

    # Somas e contagens por Chave (Plantilla x Cuenta)
    df_pg_subset[amount_col] = pd.to_numeric(df_pg_subset[amount_col], errors="coerce").fillna(0.0)
    g_pg = df_pg_subset.groupby("Chave", as_index=False).agg(
        soma_pg=(amount_col, "sum"),
        n_pg=("Chave", "size")
    )

    if "Saldo Real" not in df_cu_subset.columns:
        raise ValueError("No GL0061 processado não há a coluna 'Saldo Real'.")

    df_cu_subset["Saldo Real"] = pd.to_numeric(df_cu_subset["Saldo Real"], errors="coerce").fillna(0.0)
    g_cu = df_cu_subset.groupby("Chave", as_index=False).agg(
        soma_cu=("Saldo Real", "sum"),
        n_cu=("Chave", "size")
    )

    resumo = pd.merge(g_cu, g_pg, on="Chave", how="outer")
    resumo["soma_cu"] = pd.to_numeric(resumo["soma_cu"], errors="coerce").fillna(0.0)
    resumo["soma_pg"] = pd.to_numeric(resumo["soma_pg"], errors="coerce").fillna(0.0)
    resumo["n_cu"] = pd.to_numeric(resumo["n_cu"], errors="coerce").fillna(0).astype(int)
    resumo["n_pg"] = pd.to_numeric(resumo["n_pg"], errors="coerce").fillna(0).astype(int)
    resumo["dif"] = (resumo["soma_pg"] - resumo["soma_cu"]).round(2)
    resumo["ajustar"] = resumo["dif"].abs() > float(tol)

    # Chaves sem ajuste
    chaves_ok = set(resumo.loc[~resumo["ajustar"], "Chave"].astype(str).tolist())

    # Mantém as linhas atuais para chaves ok
    linhas_novas = []
    if not df_pg_subset.empty:
        linhas_novas.append(df_pg_subset[df_pg_subset["Chave"].isin(chaves_ok)])

    # Reconstrói chaves com diferença usando as linhas do GL
    chaves_ajustar = resumo.loc[resumo["ajustar"], "Chave"].astype(str).tolist()

    molde_por_chave = {}
    if not df_pg_subset.empty:
        molde_por_chave = df_pg_subset.groupby("Chave").head(1).set_index("Chave")

    for chave in chaves_ajustar:
        bloc_cu = df_cu_subset[df_cu_subset["Chave"] == chave]
        if bloc_cu.empty:
            continue

        molde = molde_por_chave.loc[chave].copy() if chave in molde_por_chave.index else template_base.copy()

        for _, rc in bloc_cu.iterrows():
            row = molde.copy()

            if cuenta_col is not None:
                row[cuenta_col] = rc.get("CTA", None)
            if tdate_col is not None:
                row[tdate_col] = rc.get("Fecha", None)
            if tno_col is not None:
                row[tno_col] = rc.get("Transacción", None)

            # Amount = Saldo Real do GL
            row[amount_col] = float(rc.get("Saldo Real", 0.0))

            # Recalcula Chave com mesmo formato do app
            cta_str   = _str_or_empty(row.get(cuenta_col)) if cuenta_col is not None else _str_or_empty(rc.get("CTA"))
            tdate_str = _fmt_date_ddmmyyyy(row.get(tdate_col)) if tdate_col is not None else _fmt_date_ddmmyyyy(rc.get("Fecha"))
            tno_str   = _str_or_empty(row.get(tno_col)) if tno_col is not None else _str_or_empty(rc.get("Transacción"))
            amt_str   = _fmt_num_2dec_point(row.get(amount_col))
            row["Chave"] = f"{cta_str}|{tdate_str}|{tno_str}|{amt_str}"

            try:
                row[amount_col] = float(row[amount_col])
            except Exception:
                row[amount_col] = pd.to_numeric(row[amount_col], errors="coerce")

            linhas_novas.append(pd.DataFrame([row.to_dict()]))

    bloco_cta_atualizado = pd.concat(linhas_novas, ignore_index=True) if len(linhas_novas) > 0 else pd.DataFrame(columns=df_pg.columns)
    df_pg_atualizada = pd.concat([df_pg_other, bloco_cta_atualizado], ignore_index=True)

    # Ordenação amigável se existirem colunas
    if cuenta_col is not None and tdate_col is not None:
        df_pg_atualizada = df_pg_atualizada.sort_values(by=[cuenta_col, tdate_col], ascending=[True, True], na_position="last").reset_index(drop=True)

    resumo_view = resumo[["Chave", "soma_cu", "soma_pg", "dif", "n_cu", "n_pg", "ajustar"]].sort_values(by="Chave").reset_index(drop=True)
    return df_pg_atualizada, resumo_view

# -----------------------------------------------------------------------------
# Página
# -----------------------------------------------------------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    # Botões principais (agora com CUENTA)
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

    # -------------------------------------------------------------------------
    # Modo: Estado de Cuenta (.txt)
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

                pbar.progress(70, text="Preparando visualização...")
                st.success("Arquivo processado com sucesso.")
                st.dataframe(
                    df,
                    use_container_width=True,
                    height=550,
                    column_config={c: st.column_config.NumberColumn(format="%.2f") for c in numeric_cols if c in df.columns},
                )

                pbar.progress(90, text="Gerando arquivos para download...")
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

                pbar.progress(100, text="Concluído.")
            except Exception as e:
                st.error("Erro ao processar o arquivo .txt.")
                st.exception(e)

    # -------------------------------------------------------------------------
    # Modo: Plantilla de Gastos (.xlsx/.xls)
    # -------------------------------------------------------------------------
    elif mode == "plantilla":
        upl_key_pg = st.session_state["aag_state"].setdefault("uploader_key_pg", "aag_pg_upl_1")

        st.caption("Carregue o arquivo **Excel** da *Plantilla de Gastos* (primeira aba será lida).")
        uploaded_xl = st.file_uploader(
            "Selecionar arquivo (.xlsx ou .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key=upl_key_pg,
            help="As colunas de data serão convertidas para dd/mm/aaaa e Amount formatado como número.",
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", use_container_width=True, disabled=(uploaded_xl is None))
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_pg"] = upl_key_pg + "_x"
            if "aag_plantilla_df" in st.session_state:
                del st.session_state["aag_plantilla_df"]
            st.rerun()

        if run_clicked and uploaded_xl is not None:
            pbar = st.progress(0, text="Lendo arquivo Excel...")
            try:
                name = getattr(uploaded_xl, "name", "").lower()
                engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
                df_pg = pd.read_excel(uploaded_xl, sheet_name=0, engine=engine)

                # Detecta coluna Amount (case-insensitive)
                amount_col = None
                for c in df_pg.columns:
                    if str(c).strip().lower() == "amount":
                        amount_col = c
                        break
                if amount_col is None:
                    candidates = [c for c in df_pg.columns if "amount" in str(c).strip().lower()]
                    if candidates:
                        amount_col = candidates[0]

                if amount_col is None:
                    st.error("Coluna 'Amount' não encontrada no arquivo.")
                    return

                # Datas comuns na Plantilla
                def norm(s: str) -> str:
                    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

                date_targets = {"transactiondate": None, "due_date": None, "invoicedate": None, "invoice_date": None}
                found_date_cols = []
                for c in df_pg.columns:
                    nc = norm(c)
                    if nc == "transactiondate":
                        date_targets["transactiondate"] = c
                    elif nc in ("duedate", "due_date"):
                        date_targets["due_date"] = c
                    elif nc in ("invoicedate", "invoice_date"):
                        if date_targets.get("invoicedate") is None:
                            date_targets["invoicedate"] = c

                for key, col in date_targets.items():
                    if col and col in df_pg.columns:
                        df_pg[col] = pd.to_datetime(df_pg[col], errors="coerce", dayfirst=True).dt.date
                        found_date_cols.append(col)

                # Numérico em Amount
                df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

                # Cria 'Chave'
                def _find_col_ci_local(df: pd.DataFrame, targets: list[str]):
                    cols_map = {re.sub(r"[^a-z0-9]", "", str(c).lower()): c for c in df.columns}
                    for t in targets:
                        key = re.sub(r"[^a-z0-9]", "", t.lower())
                        if key in cols_map:
                            return cols_map[key]
                    return None

                cuenta_col   = _find_col_ci_local(df_pg, ["Cuenta"])
                tdate_col    = _find_col_ci_local(df_pg, ["TransactionDate", "Transaction Date", "TransDate"])
                tno_col      = _find_col_ci_local(df_pg, ["TransactionNo", "Transaction No", "TransNo", "Transaction_Number"])
                amount_col_ci = amount_col

                tdate_str = df_pg[tdate_col].apply(_fmt_date_ddmmyyyy) if tdate_col else ""
                tno_str   = df_pg[tno_col].apply(_str_or_empty) if tno_col else ""
                cuenta_str= df_pg[cuenta_col].apply(_str_or_empty) if cuenta_col else ""
                amount_str= df_pg[amount_col_ci].apply(_fmt_num_2dec_point) if amount_col_ci else ""

                df_pg["Chave"] = (
                    (cuenta_str if isinstance(cuenta_str, pd.Series) else pd.Series([""]*len(df_pg))) + "|" +
                    (tdate_str  if isinstance(tdate_str,  pd.Series) else pd.Series([""]*len(df_pg))) + "|" +
                    (tno_str    if isinstance(tno_str,    pd.Series) else pd.Series([""]*len(df_pg))) + "|" +
                    (amount_str if isinstance(amount_str, pd.Series) else pd.Series([""]*len(df_pg)])
                )

                st.session_state["aag_plantilla_df"] = df_pg.copy()

                pbar.progress(70, text="Preparando visualização...")
                st.success("Arquivo carregado com sucesso.")

                col_cfg = {str(amount_col): st.column_config.NumberColumn(format="%.2f")}
                for dc in found_date_cols:
                    col_cfg[dc] = st.column_config.DateColumn(format="DD/MM/YYYY")

                st.dataframe(df_pg, use_container_width=True, height=550, column_config=col_cfg)

                pbar.progress(90, text="Gerando arquivos para download...")
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        "Baixar CSV (Plantilla Gastos)",
                        df_pg.to_csv(index=False).encode("utf-8"),
                        "plantilla_gastos.csv",
                        "text/csv",
                        use_container_width=True,
                    )
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes_format(
                        df_pg, sheet_name="PlantillaGastos", numeric_cols=[amount_col], date_cols=found_date_cols
                    )
                    st.download_button(
                        "Baixar XLSX (Plantilla Gastos)",
                        xlsx_bytes,
                        "plantilla_gastos.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                pbar.progress(100, text="Concluído.")
            except Exception as e:
                st.error("Erro ao processar o arquivo Excel.")
                st.exception(e)

    # -------------------------------------------------------------------------
    # Modo: Analise — compara Estado de Cuenta x Plantilla
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
                st.metric("Soma Estado de Cuentas", f"{df_ec_agg['Saldo_Estado_Cuenta'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            with c3:
                st.metric("Soma Plantilla de Gastos", f"{df_pg_agg['Saldo_Plantilla_Gastos'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

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
                    "Baixar CSV (Analise)",
                    df_show.to_csv(index=False).encode("utf-8"),
                    "analise_contas.csv",
                    "text/csv",
                    use_container_width=True,
                )
            with col_d2:
                xlsx_bytes = to_xlsx_bytes_format(
                    df_show, "Analise",
                    numeric_cols=["Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos", "Diferença"],
                    date_cols=[],
                )
                st.download_button(
                    "Baixar XLSX (Analise)",
                    xlsx_bytes,
                    "analise_contas.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            pbar.progress(100, text="Concluído.")
        except Exception as e:
            st.error("Erro durante a comparação.")
            st.exception(e)

    # -------------------------------------------------------------------------
    # Modo: Cuenta (GL0061) — corrigido para persistir UI e botão sempre visível
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

        # Limpar: reseta uploader e remove DF da sessão
        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_cuenta"] = upl_key + "_x"
            if "aag_cuenta_df" in st.session_state:
                del st.session_state["aag_cuenta_df"]
            st.rerun()

        # Processar novo upload (se houver)
        if run_clicked and uploaded is not None:
            raw = uploaded.getvalue()
            try:
                text = raw.decode("utf-8")
            except Exception:
                text = raw.decode("latin-1")

            try:
                df_new = parse_cuenta_gl(text)
            except Exception as e:
                st.error("Erro ao interpretar o arquivo GL0061.")
                st.exception(e)
                return

            if df_new.empty:
                st.error("Nenhuma linha reconhecida no arquivo GL0061.")
                return

            # Monta Chave
            cta_str   = df_new["CTA"].apply(_str_or_empty) if "CTA" in df_new.columns else pd.Series([""]*len(df_new))
            fecha_str = df_new["Fecha"].apply(_fmt_date_ddmmyyyy) if "Fecha" in df_new.columns else pd.Series([""]*len(df_new))
            tran_str  = df_new["Transacción"].apply(_str_or_empty) if "Transacción" in df_new.columns else pd.Series([""]*len(df_new))
            sreal_str = df_new["Saldo Real"].apply(_fmt_num_2dec_point) if "Saldo Real" in df_new.columns else pd.Series([""]*len(df_new))
            df_new["Chave"] = cta_str + "|" + fecha_str + "|" + tran_str + "|" + sreal_str

            st.session_state["aag_cuenta_df"] = df_new.copy()
            st.success("Cuenta carregada e processada com sucesso.")

        # A partir daqui, renderiza SEMPRE que houver DF na sessão, mesmo sem upload
        df_cu = st.session_state.get("aag_cuenta_df")

        if df_cu is not None and not df_cu.empty:
            # Configura visualização
            date_cols = [c for c in ["Fecha", "Fechado"] if c in df_cu.columns]
            col_cfg = {
                "Debe": st.column_config.NumberColumn(format="%.2f"),
                "Haber": st.column_config.NumberColumn(format="%.2f"),
                "Saldo Real": st.column_config.NumberColumn(format="%.2f"),
                "Saldo": st.column_config.NumberColumn(format="%.2f"),
            }
            for dc in date_cols:
                col_cfg[dc] = st.column_config.DateColumn(format="DD/MM/YYYY")

            st.dataframe(df_cu, use_container_width=True, height=580, column_config=col_cfg)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "Baixar CSV (Cuenta)",
                    df_cu.to_csv(index=False).encode("utf-8"),
                    "cuenta.csv",
                    "text/csv",
                    use_container_width=True
                )
            with col2:
                xlsx_bytes = to_xlsx_bytes_format(
                    df_cu, "Cuenta",
                    numeric_cols=["Debe","Haber","Saldo Real","Saldo"] if all(c in df_cu.columns for c in ["Debe","Haber","Saldo Real","Saldo"]) else ["Saldo Real","Saldo"],
                    date_cols=date_cols
                )
                st.download_button(
                    "Baixar XLSX (Cuenta)",
                    xlsx_bytes,
                    "cuenta.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # ----- Botão fixo: Plantilla Gastos Atualizada -----
            df_pg = st.session_state.get("aag_plantilla_df")
            st.divider()
            st.subheader("🧹 Plantilla Gastos Atualizada (baseada na Cuenta carregada)")

            if df_pg is None or df_pg.empty:
                st.info("Carregue a **Plantilla de Gastos** na aba correspondente para habilitar a atualização.")
                return

            tol_update = st.number_input(
                "Tolerância para comparar saldos (|Plantilla - Cuenta|)",
                min_value=0.00, value=0.01, step=0.01
            )

            gen_clicked = st.button("🧩 Gerar Plantilla Gastos Atualizada", type="primary", use_container_width=True)
            if gen_clicked:
                try:
                    df_pg_atual, resumo_adj = build_plantilla_atualizada_com_cuenta(df_pg.copy(), df_cu.copy(), tol=tol_update)
                    st.session_state["aag_plantilla_atualizada_df"] = df_pg_atual

                    st.success("Plantilla de Gastos atualizada com sucesso para a conta carregada.")
                    st.caption("Resumo das chaves comparadas (apenas desta CTA):")
                    st.dataframe(
                        resumo_adj,
                        use_container_width=True,
                        height=320,
                        column_config={
                            "soma_cu": st.column_config.NumberColumn(format="%.2f"),
                            "soma_pg": st.column_config.NumberColumn(format="%.2f"),
                            "dif": st.column_config.NumberColumn(format="%.2f"),
                        }
                    )

                    # Detecta colunas de data para exportação
                    amount_col_exp = _find_col_ci(df_pg_atual, ["Amount"])
                    date_cols_exp = []
                    for cand in ["TransactionDate", "Transaction Date", "TransDate", "Due_Date", "Invoice_Date", "DueDate", "InvoiceDate"]:
                        c_real = _find_col_ci(df_pg_atual, [cand])
                        if c_real is not None:
                            date_cols_exp.append(c_real)
                    date_cols_exp = list(dict.fromkeys(date_cols_exp))

                    st.subheader("⬇️ Baixar Plantilla Gastos Atualizada")
                    col_u1, col_u2 = st.columns(2)
                    with col_u1:
                        st.download_button(
                            "Baixar CSV (Plantilla Atualizada)",
                            df_pg_atual.to_csv(index=False).encode("utf-8"),
                            "plantilla_gastos_atualizada.csv",
                            "text/csv",
                            use_container_width=True
                        )
                    with col_u2:
                        xlsx_upd = to_xlsx_bytes_format(
                            df_pg_atual,
                            sheet_name="PlantillaAtualizada",
                            numeric_cols=[amount_col_exp] if amount_col_exp else [],
                            date_cols=date_cols_exp
                        )
                        st.download_button(
                            "Baixar XLSX (Plantilla Atualizada)",
                            xlsx_upd,
                            "plantilla_gastos_atualizada.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                except Exception as e:
                    st.error("Falha ao gerar Plantilla Gastos Atualizada.")
                    st.exception(e)
        else:
            st.info("Carregue um arquivo **GL0061** ou mantenha o existente para gerar a Plantilla atualizada.")

    else:
        st.info("Selecione um modo acima para continuar.")

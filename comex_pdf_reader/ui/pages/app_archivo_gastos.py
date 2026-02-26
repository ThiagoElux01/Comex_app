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
    """
    Converte strings como '12,345.67-' em float (negativo).
    """
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

    # ==== Helpers globais extras ====
    def _norm_conta(x) -> str:
        """Normaliza o número da conta: só dígitos e remove zeros à esquerda."""
        s = re.sub(r"\D", "", str(x))
        s = s.lstrip("0")
        return s if s else ""
    
    def _split_chave(ch: str):
        """
        Divide 'Chave' no formato esperado:
            Plantilla: Cuenta|TransactionDate|TransactionNo|Amount
            Cuenta(GL0061): CTA|Fecha|Transacción|Saldo Real
        Retorna: (cuenta_str, date(date|None), trans_str, amount(float|None))
        """
        parts = (str(ch) if ch is not None else "").split("|")
        while len(parts) < 4:
            parts.append("")
        cuenta_str = parts[0].strip()
    
        date_str = parts[1].strip()
        date_val = pd.to_datetime(date_str, errors="coerce", dayfirst=True)
        date_val = date_val.date() if pd.notna(date_val) else None
    
        trans_str = parts[2].strip()
        amount_str = parts[3].strip()
        amount_val = _clean_num(amount_str)
    
        return cuenta_str, date_val, trans_str, amount_val
    
    def _find_col_ci_generic(df: pd.DataFrame, targets: list[str]):
        """
        Busca coluna por nome, ignorando maiúsc./minúsc. e sinais. Retorna o primeiro match.
        """
        if df is None or df.empty:
            return None
        mm = {re.sub(r"[^a-z0-9]", "", str(c).lower()): c for c in df.columns}
        for t in targets:
            k = re.sub(r"[^a-z0-9]", "", t.lower())
            if k in mm:
                return mm[k]
        return None

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

    # === Ajuste de data no dataframe de cuentas ===
    # Se houver coluna 'Fechado' (alguns dumps usam esse nome), trata também.
    for date_col in ["Fecha", "Fechado"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.date

    return df

# ----------------------------------------------------------------------
# Limpeza da Plantilla de Gastos orientada pela Cuenta (GL0061)
# ----------------------------------------------------------------------
def _ensure_pg_chave(df_pg: pd.DataFrame) -> pd.DataFrame:
    """
    Garante a coluna 'Chave' na Plantilla (se não existir, tenta criar).
    Usa as mesmas regras da aba Plantilla.
    """
    if df_pg is None or df_pg.empty:
        return df_pg

    if "Chave" in df_pg.columns:
        return df_pg

    # Detecta colunas relevantes
    cuenta_col = _find_col_ci_generic(df_pg, ["Cuenta"])
    amount_col = _find_col_ci_generic(df_pg, ["Amount"])
    tdate_col  = _find_col_ci_generic(df_pg, ["TransactionDate", "Transaction Date", "TransDate"])
    tno_col    = _find_col_ci_generic(df_pg, ["TransactionNo", "Transaction No", "TransNo", "Transaction_Number"])

    # Se não existir Amount ou Cuenta, não há como formar a Chave
    if amount_col is None or cuenta_col is None:
        return df_pg

    # Formata componentes
    tdate_str   = df_pg[tdate_col].apply(_fmt_date_ddmmyyyy) if tdate_col else pd.Series([""] * len(df_pg))
    tno_str     = df_pg[tno_col].apply(_str_or_empty) if tno_col else pd.Series([""] * len(df_pg))
    cuenta_str  = df_pg[cuenta_col].apply(_str_or_empty)
    df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")
    amount_str  = df_pg[amount_col].apply(_fmt_num_2dec_point)

    df_pg["Chave"] = cuenta_str + "|" + tdate_str + "|" + tno_str + "|" + amount_str
    return df_pg


def _build_pg_rows_from_cuenta(df_cuenta_cta: pd.DataFrame,
                               pg_columns: list[str],
                               cuenta_col: str | None,
                               amount_col: str | None,
                               tdate_col: str | None,
                               tno_col: str | None) -> pd.DataFrame:
    """
    Cria linhas de Plantilla a partir das linhas de Cuenta (1‑para‑1 com a 'Chave' de Cuenta).
    Mantém todas as colunas da Plantilla (preenche desconhecidas com None).
    """
    rows = []
    for _, r in df_cuenta_cta.iterrows():
        chave = r.get("Chave", "")
        cta_str, fdate, trans, amt = _split_chave(chave)

        new_row = {c: None for c in pg_columns}
        if "Chave" in pg_columns:
            new_row["Chave"] = chave
        if cuenta_col:
            new_row[cuenta_col] = cta_str
        if amount_col:
            new_row[amount_col] = float(amt) if amt is not None else None
        if tdate_col:
            new_row[tdate_col] = fdate
        if tno_col:
            new_row[tno_col] = trans

        rows.append(new_row)

    return pd.DataFrame(rows, columns=pg_columns)


    def clean_plantilla_by_cuenta(df_cuenta: pd.DataFrame,
                                  df_pg: pd.DataFrame,
                                  tol: float = 0.0,
                                  only_problem_keys: bool = True):
        """
        Limpa a Plantilla de Gastos (apenas da CTA carregada em Cuenta), usando as chaves da Cuenta como referência.
    
        Regra:
          - Calcula (por Chave) soma e contagem em Cuenta (Saldo Real) e Plantilla (Amount).
          - Se |dif| <= tol e contagens iguais -> mantém Plantilla como está para essa Chave.
          - Caso contrário -> reconstrói as linhas dessa Chave na Plantilla com base nas linhas da Cuenta
            (mesmo número de linhas e mesma Chave; consequentemente, Amount segue o Saldo Real).
    
        Parâmetros
          df_cuenta          : DataFrame retornado do GL0061 (com 'Chave', 'CTA', 'Saldo Real', 'Fecha', 'Transacción')
          df_pg              : DataFrame atual da Plantilla (de preferência o já carregado na aba Plantilla)
          tol                : tolerância de diferença absoluta entre somas (default 0.00)
          only_problem_keys  : True -> reconstrói apenas chaves com divergência (ou contagem diferente)
                               False -> reconstrói TODAS as chaves da CTA carregada (modo “forçar total”)
    
        Retorna
          (df_pg_clean, resumo_dict)
        """
        if df_cuenta is None or df_cuenta.empty:
            raise ValueError("Cuenta (GL0061) inexistente ou vazia.")
        if df_pg is None or df_pg.empty:
            raise ValueError("Plantilla de Gastos inexistente ou vazia.")
    
        if "CTA" not in df_cuenta.columns or "Chave" not in df_cuenta.columns:
            raise ValueError("Cuenta precisa conter as colunas 'CTA' e 'Chave'.")
        # Pega a CTA do arquivo (GL0061 é por conta)
        cta_raw = str(df_cuenta["CTA"].dropna().astype(str).iloc[0])
        cta_norm = _norm_conta(cta_raw)
    
        # Garante 'Chave' na Plantilla
        df_pg = _ensure_pg_chave(df_pg.copy())
    
        # Detecta colunas pivot da Plantilla
        cuenta_col = _find_col_ci_generic(df_pg, ["Cuenta"])
        amount_col = _find_col_ci_generic(df_pg, ["Amount"])
        tdate_col  = _find_col_ci_generic(df_pg, ["TransactionDate", "Transaction Date", "TransDate"])
        tno_col    = _find_col_ci_generic(df_pg, ["TransactionNo", "Transaction No", "TransNo", "Transaction_Number"])
    
        if cuenta_col is None or amount_col is None:
            raise ValueError("Plantilla precisa conter as colunas 'Cuenta' e 'Amount'.")
    
        # Normaliza CTA na Plantilla (a partir da própria 'Chave' para garantir coerência)
        if "Chave" in df_pg.columns:
            pg_cuenta = df_pg["Chave"].astype(str).str.split("|").str[0]
        else:
            pg_cuenta = df_pg[cuenta_col].astype(str)
        df_pg["__cta_norm__"] = pg_cuenta.apply(_norm_conta)
    
        # Filtra somente a CTA carregada
        df_pg_cta = df_pg[df_pg["__cta_norm__"] == cta_norm].copy()
        df_pg_others = df_pg[df_pg["__cta_norm__"] != cta_norm].copy()
    
        # --- Stats por Chave ---
        # Cuenta
        df_cuenta_cta = df_cuenta.copy()
        df_cuenta_cta["__cta_norm__"] = df_cuenta_cta["CTA"].apply(_norm_conta)
        df_cuenta_cta = df_cuenta_cta[df_cuenta_cta["__cta_norm__"] == cta_norm]
        if df_cuenta_cta.empty:
            raise ValueError("A CTA do arquivo Cuenta não foi identificada nas linhas lidas.")
    
        if "Saldo Real" not in df_cuenta_cta.columns:
            raise ValueError("Cuenta deve conter a coluna 'Saldo Real'.")
    
        grp_cuenta = df_cuenta_cta.groupby("Chave", as_index=False).agg(
            cnt_cuenta=("Chave", "size"),
            sum_cuenta=("Saldo Real", "sum")
        )
    
        # Plantilla
        df_pg_cta[amount_col] = pd.to_numeric(df_pg_cta[amount_col], errors="coerce").fillna(0.0)
        grp_pg = df_pg_cta.groupby("Chave", as_index=False).agg(
            cnt_pg=("Chave", "size"),
            sum_pg=(amount_col, "sum")
        )
    
        # Merge e definição de chaves problemáticas
        cmp_keys = pd.merge(grp_cuenta, grp_pg, on="Chave", how="outer")
        cmp_keys["cnt_cuenta"] = pd.to_numeric(cmp_keys["cnt_cuenta"], errors="coerce").fillna(0).astype(int)
        cmp_keys["cnt_pg"]     = pd.to_numeric(cmp_keys["cnt_pg"], errors="coerce").fillna(0).astype(int)
        cmp_keys["sum_cuenta"] = pd.to_numeric(cmp_keys["sum_cuenta"], errors="coerce").fillna(0.0)
        cmp_keys["sum_pg"]     = pd.to_numeric(cmp_keys["sum_pg"], errors="coerce").fillna(0.0)
        cmp_keys["abs_diff"]   = (cmp_keys["sum_pg"] - cmp_keys["sum_cuenta"]).abs()
    
        # Chaves presentes apenas em Cuenta ou com diferença de soma ou contagem diferente são "problemáticas"
        prob_mask = (cmp_keys["abs_diff"] > float(tol)) | (cmp_keys["cnt_cuenta"] != cmp_keys["cnt_pg"]) | cmp_keys["sum_pg"].isna()
        if only_problem_keys:
            keys_to_fix = set(cmp_keys.loc[prob_mask, "Chave"].dropna().astype(str))
        else:
            keys_to_fix = set(cmp_keys["Chave"].dropna().astype(str))
    
        # --- Reconstrução ---
        # a) Mantém linhas da Plantilla da CTA que NÃO precisam de ajuste
        keep_mask = ~df_pg_cta["Chave"].astype(str).isin(keys_to_fix)
        df_pg_keep = df_pg_cta[keep_mask].copy()
    
        # b) Constrói linhas novas a partir da Cuenta para as chaves que precisam de ajuste
        if keys_to_fix:
            df_cuenta_fix = df_cuenta_cta[df_cuenta_cta["Chave"].astype(str).isin(keys_to_fix)].copy()
            df_pg_new = _build_pg_rows_from_cuenta(
                df_cuenta_fix,
                pg_columns=list(df_pg.columns),  # preserva estrutura/ordem
                conta_col=cuenta_col if (cuenta_col in df_pg.columns) else None,
                amount_col=amount_col if (amount_col in df_pg.columns) else None,
                tdate_col=tdate_col if (tdate_col in df_pg.columns) else None,
                tno_col=tno_col if (tno_col in df_pg.columns) else None,
            )
        else:
            df_pg_new = df_pg_keep.iloc[0:0].copy()
    
        # c) Recompõe Plantilla: outros CTAs + (CTA corrente: keep + new)
        df_pg_clean = pd.concat([df_pg_others, df_pg_keep, df_pg_new], ignore_index=True)
        if "__cta_norm__" in df_pg_clean.columns:
            df_pg_clean = df_pg_clean.drop(columns=["__cta_norm__"])
    
        # Resumo
        resumo = {
            "cta": cta_norm,
            "total_pg_cta_antes": int(len(df_pg_cta)),
            "total_pg_cta_depois": int(len(df_pg_keep) + len(df_pg_new)),
            "chaves_total_cta": int(cmp_keys["Chave"].notna().sum()),
            "chaves_ajustadas": int(len(keys_to_fix)),
            "tol": float(tol),
            "only_problem_keys": bool(only_problem_keys),
        }
        return df_pg_clean, resumo
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
        # Para manter tipo data no Excel, converte objetos date para datetime64[ns] (sem horário)
        df_to_save = df.copy()
        for dc in date_cols:
            if dc in df_to_save.columns:
                # Se a coluna está como datetime.date, converte para datetime64
                df_to_save[dc] = pd.to_datetime(df_to_save[dc], errors="coerce")

        df_to_save.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        # Estilos de cabeçalho
        BLUE = "FF0077B6"
        WHITE = "FFFFFFFF"
        fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
        font_white_bold = Font(color=WHITE, bold=True)
        for cell in ws[1]:
            cell.fill = fill_blue
            cell.font = font_white_bold

        # Formatação numérica
        for col_name in numeric_cols:
            if col_name not in df_to_save.columns:
                continue
            col_idx = df_to_save.columns.get_loc(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)) and cell.value is not None:
                    cell.number_format = '#,##0.00'

        # Formatação de data
        for col_name in date_cols:
            if col_name not in df_to_save.columns:
                continue
            col_idx = df_to_save.columns.get_loc(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                # openpyxl trata datetime como Python datetime; só aplicar formato
                if cell.value:
                    cell.number_format = 'dd/mm/yyyy'

        # Ajuste de largura
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

                # Salva o DF base (sem a linha TOTAL) para uso na aba Analise
                st.session_state["aag_estado_df"] = df_base.copy()

                # ======== LINHA TOTAL ========
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
                    column_config={
                        c: st.column_config.NumberColumn(format="%.2f") for c in numeric_cols if c in df.columns
                    },
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

                # === Datas na Plantilla: TransactionDate, Due_Date, Invoice_Date ===
                # normalizador de nomes para encontrar colunas mesmo com variações
                def norm(s: str) -> str:
                    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

                date_targets = {
                    "transactiondate": None,
                    "due_date": None,
                    "invoicedate": None,
                    "invoice_date": None,
                }
                # Mapear colunas do DF a alvos
                found_date_cols = []
                for c in df_pg.columns:
                    nc = norm(c)
                    if nc == "transactiondate":
                        date_targets["transactiondate"] = c
                    elif nc in ("duedate", "due_date"):
                        date_targets["due_date"] = c
                    elif nc in ("invoicedate", "invoice_date"):
                        # Prioriza a primeira encontrada
                        if date_targets.get("invoicedate") is None:
                            date_targets["invoicedate"] = c

                # Converte cada coluna encontrada para data (dd/mm/aaaa)
                for key, col in date_targets.items():
                    if col and col in df_pg.columns:
                        df_pg[col] = pd.to_datetime(df_pg[col], errors="coerce", dayfirst=True).dt.date
                        found_date_cols.append(col)

                # Garante tipo numérico em Amount
                df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

                # === Criar coluna 'Chave' na Plantilla de Gastos ===
                # Detecta campos necessários (case-insensitive, tolerando variações)
                def _find_col_ci(df: pd.DataFrame, targets: list[str]):
                    cols_map = {re.sub(r"[^a-z0-9]", "", str(c).lower()): c for c in df.columns}
                    for t in targets:
                        key = re.sub(r"[^a-z0-9]", "", t.lower())
                        if key in cols_map:
                            return cols_map[key]
                    return None
                
                cuenta_col   = _find_col_ci(df_pg, ["Cuenta"])
                tdate_col    = _find_col_ci(df_pg, ["TransactionDate", "Transaction Date", "TransDate"])
                tno_col      = _find_col_ci(df_pg, ["TransactionNo", "Transaction No", "TransNo", "Transaction_Number"])
                amount_col_ci = amount_col  # já detectado acima
                
                # Converter datas detectadas para dd/mm/aaaa (string) e números para 2 casas
                tdate_str = df_pg[tdate_col].apply(_fmt_date_ddmmyyyy) if tdate_col else ""
                tno_str   = df_pg[tno_col].apply(_str_or_empty) if tno_col else ""
                cuenta_str= df_pg[cuenta_col].apply(_str_or_empty) if cuenta_col else ""
                amount_str= df_pg[amount_col_ci].apply(_fmt_num_2dec_point) if amount_col_ci else ""
                
                # Concatena com pipe
                df_pg["Chave"] = (
                    (cuenta_str if isinstance(cuenta_str, pd.Series) else pd.Series([""]*len(df_pg))) + "|" +
                    (tdate_str  if isinstance(tdate_str,  pd.Series) else pd.Series([""]*len(df_pg))) + "|" +
                    (tno_str    if isinstance(tno_str,    pd.Series) else pd.Series([""]*len(df_pg))) + "|" +
                    (amount_str if isinstance(amount_str, pd.Series) else pd.Series([""]*len(df_pg)))
                )
                
                # Salva o DF para uso na aba Analise (agora com 'Chave')
                st.session_state["aag_plantilla_df"] = df_pg.copy()

                pbar.progress(70, text="Preparando visualização...")
                st.success("Arquivo carregado com sucesso.")

                # Configuração de colunas para visualização
                col_cfg = {
                    str(amount_col): st.column_config.NumberColumn(format="%.2f")
                }
                for dc in found_date_cols:
                    col_cfg[dc] = st.column_config.DateColumn(format="DD/MM/YYYY")

                st.dataframe(
                    df_pg,
                    use_container_width=True,
                    height=550,
                    column_config=col_cfg,
                )

                pbar.progress(90, text="Gerando arquivos para download...")
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        label="Baixar CSV (Plantilla Gastos)",
                        data=df_pg.to_csv(index=False).encode("utf-8"),
                        file_name="plantilla_gastos.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes_format(
                        df_pg,
                        sheet_name="PlantillaGastos",
                        numeric_cols=[amount_col],
                        date_cols=found_date_cols,
                    )
                    st.download_button(
                        label="Baixar XLSX (Plantilla Gastos)",
                        data=xlsx_bytes,
                        file_name="plantilla_gastos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                pbar.progress(100, text="Concluído.")
            except Exception as e:
                st.error("Erro ao processar o arquivo Excel.")
                st.exception(e)

    # -------------------------------------------------------------------------
    # Modo: Analise — compara Estado de Cuenta (CTA/Período) x Plantilla (Cuenta/Amount)
    # -------------------------------------------------------------------------
    elif mode == "asientos":  # Analise
        st.subheader("🔍 Analise: Estado de Cuenta x Plantilla de Gastos")

        # Recupera datasets da sessão
        df_ec = st.session_state.get("aag_estado_df", None)       # Estado de Cuenta (sem linha TOTAL)
        df_pg = st.session_state.get("aag_plantilla_df", None)    # Plantilla de Gastos (primeira aba)

        # Checagens
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

        # Parâmetros
        tol = st.number_input("Valor de Tolerância", min_value=0.00, value=0.01, step=0.01)

        # Helpers
        def _norm_conta(x) -> str:
            """Normaliza o número da conta: apenas dígitos; remove zeros à esquerda."""
            s = re.sub(r"\D", "", str(x))
            s = s.lstrip("0")
            return s if s else ""

        pbar = st.progress(0, text="Consolidando Estado de Cuenta...")

        # -----------------------
        # 1) Estado de Cuenta (CTA/Período)
        # -----------------------
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

        # -----------------------
        # 2) Plantilla de Gastos (Cuenta/Amount)
        # -----------------------
        try:
            # Detecta colunas 'Cuenta' e 'Amount' (case-insensitive)
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

        # -----------------------
        # 3) Comparação
        # -----------------------
        try:
            df_cmp = pd.merge(df_ec_agg, df_pg_agg, on="Cuenta", how="outer")
            for c in ["Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos"]:
                df_cmp[c] = pd.to_numeric(df_cmp[c], errors="coerce").fillna(0.0)

            df_cmp["Diferença"] = (df_cmp["Saldo_Plantilla_Gastos"] - df_cmp["Saldo_Estado_Cuenta"]).round(2)

            # Flag só para filtro (não exibida)
            df_cmp["_div"] = df_cmp["Diferença"].abs() > float(tol)

            # Ordenação por número da conta (ordem numérica crescente)
            df_cmp["_cuenta_num"] = pd.to_numeric(df_cmp["Cuenta"], errors="coerce")
            df_cmp = df_cmp.sort_values(by="_cuenta_num", ascending=True).drop(columns=["_cuenta_num"]).reset_index(drop=True)

            pbar.progress(90, text="Preparando visualização...")

            # Métricas
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

            # Filtro: apenas divergentes
            only_div = st.checkbox("Mostrar apenas contas com divergência", value=True)
            df_show = df_cmp[df_cmp["_div"]] if only_div else df_cmp

            # Exibição sem coluna de controle
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

            # Downloads (já ordenados pela Cuenta)
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
    # Modo: Cuenta (GL0061)
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

            # Campos: CTA, Fecha, Transacción, Saldo Real
            cta_str   = df["CTA"].apply(_str_or_empty) if "CTA" in df.columns else pd.Series([""]*len(df))
            fecha_str = df["Fecha"].apply(_fmt_date_ddmmyyyy) if "Fecha" in df.columns else pd.Series([""]*len(df))
            tran_str  = df["Transacción"].apply(_str_or_empty) if "Transacción" in df.columns else pd.Series([""]*len(df))
            sreal_str = df["Saldo Real"].apply(_fmt_num_2dec_point) if "Saldo Real" in df.columns else pd.Series([""]*len(df))
            
            df["Chave"] = cta_str + "|" + fecha_str + "|" + tran_str + "|" + sreal_str
            
            # Salva com 'Chave'
            st.session_state["aag_cuenta_df"] = df.copy()

            # Colunas de data para exibição (Fecha e/ou Fechado)
            date_cols = [c for c in ["Fecha", "Fechado"] if c in df.columns]
            col_cfg = {
                "Debe": st.column_config.NumberColumn(format="%.2f"),
                "Haber": st.column_config.NumberColumn(format="%.2f"),
                "Saldo Real": st.column_config.NumberColumn(format="%.2f"),
                "Saldo": st.column_config.NumberColumn(format="%.2f"),
            }
            for dc in date_cols:
                col_cfg[dc] = st.column_config.DateColumn(format="DD/MM/YYYY")

            st.dataframe(
                df,
                use_container_width=True,
                height=600,
                column_config=col_cfg,
            )

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
            # ============================================================
            # Limpeza da Plantilla com base na Cuenta carregada (CTA atual)
            # ============================================================
            st.markdown("---")
            st.subheader("🧹 Limpar Plantilla de Gastos (somente CTA do arquivo carregado)")
            if "aag_plantilla_df" not in st.session_state or st.session_state["aag_plantilla_df"] is None:
                st.info("Carregue primeiro a **Plantilla de Gastos** na aba correspondente para habilitar esta função.")
            else:
                df_pg_atual = st.session_state["aag_plantilla_df"]
                with st.expander("Opções de processamento", expanded=True):
                    tol_pg = st.number_input("Tolerância para aceitar a soma por 'Chave' (em valor absoluto)", min_value=0.00, value=0.00, step=0.01)
                    only_prob = st.checkbox("Reconstruir apenas chaves com divergência (recomendado)", value=True)
                    force_all = st.checkbox("Forçar reconstrução **total** da CTA (substitui todas as chaves desta CTA)", value=False)
                    if force_all:
                        only_prob = False
                col_a, col_b = st.columns([2, 1])
                with col_a:
                    do_clean = st.button("🔧 Executar limpeza da Plantilla para a CTA atual", type="primary", use_container_width=True)
                with col_b:
                    replace_session = st.checkbox("Substituir Plantilla em memória pela versão limpa", value=False)

                if do_clean:
                    try:
                        df_pg_clean, resumo = clean_plantilla_by_cuenta(
                            df_cuenta=df,
                            df_pg=df_pg_atual,
                            tol=tol_pg,
                            only_problem_keys=only_prob
                        )

                        st.success("Plantilla limpa gerada com sucesso.")
                        st.write("**Resumo da CTA processada:**")
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.metric("CTA (normalizada)", resumo["cta"])
                        with c2:
                            st.metric("Linhas antes (Plantilla, CTA)", resumo["total_pg_cta_antes"])
                        with c3:
                            st.metric("Linhas depois (Plantilla, CTA)", resumo["total_pg_cta_depois"])
                        c4, c5 = st.columns(2)
                        with c4:
                            st.metric("Chaves (CTA)", resumo["chaves_total_cta"])
                        with c5:
                            st.metric("Chaves ajustadas", resumo["chaves_ajustadas"])

                        # Prévia
                        st.dataframe(
                            df_pg_clean.head(100),
                            use_container_width=True,
                            height=400,
                        )

                        # Downloads
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            st.download_button(
                                "Baixar CSV (Plantilla LIMPA - CTA atual)",
                                df_pg_clean.to_csv(index=False).encode("utf-8"),
                                "plantilla_gastos_limpia.csv",
                                "text/csv",
                                use_container_width=True
                            )
                        with col_d2:
                            xlsx_bytes_pg = to_xlsx_bytes_format(
                                df_pg_clean,
                                sheet_name="PlantillaLIMPA",
                                numeric_cols=[_find_col_ci_generic(df_pg_clean, ["Amount"]) or "Amount"],
                                date_cols=[c for c in df_pg_clean.columns if str(c).lower() in {"transactiondate","due_date","invoice_date","invoicedate"}]
                            )
                            st.download_button(
                                "Baixar XLSX (Plantilla LIMPA - CTA atual)",
                                xlsx_bytes_pg,
                                "plantilla_gastos_limpia.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                        # Atualiza sessão se solicitado
                        if replace_session:
                            st.session_state["aag_plantilla_df"] = df_pg_clean.copy()
                            st.info("A Plantilla em memória foi substituída pela versão LIMPA (CTA atual). Agora sua aba Analise refletirá a limpeza.")
                    except Exception as e:
                        st.error("Falha ao limpar a Plantilla com base na Cuenta.")
                        st.exception(e)

    else:
        st.info("Selecione um modo acima para continuar.")

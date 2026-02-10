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

    # Última ação (reserva)
    aag.setdefault("last_action", None)

    # Modo atual da página
    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"  # default

def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode

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

# -----------------------------------------------------------------------------
# Export XLSX com máscara numérica #,##0.00 (mantém tipo numérico)
# -----------------------------------------------------------------------------
def to_xlsx_bytes_numformat(df: pd.DataFrame, sheet_name: str, numeric_cols: list[str]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        # Aplica máscara #,##0.00 nas colunas numéricas indicadas
        for col_name in numeric_cols:
            if col_name not in df.columns:
                continue
            col_idx = df.columns.get_loc(col_name) + 1  # 1-based
            for row in range(2, ws.max_row + 1):  # pulando cabeçalho
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)) and cell.value is not None:
                    cell.number_format = '#,##0.00'

        # Cabeçalho estilizado (azul) e ajuste de largura simples
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

# -----------------------------------------------------------------------------
# Página
# -----------------------------------------------------------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    # Botões principais
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        if st.button("Estado de Cuenta", use_container_width=True):
            _set_mode("estado")
    with col_b2:
        if st.button("Plantilla Gastos", use_container_width=True):
            _set_mode("plantilla")
    with col_b3:
        # RÓTULO ATUALIZADO
        if st.button("Analise", use_container_width=True):
            _set_mode("asientos")

    mode = st.session_state["aag_mode"]
    st.divider()

    # -------------------------------------------------------------------------
    # Modo: Estado de Cuenta (.txt) — com totalizador (sem Styler)
    # -------------------------------------------------------------------------
    if mode == "estado":
        # Garante que a key do uploader exista
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
            # Força reset do uploader
            st.session_state["aag_state"]["uploader_key_estado"] = upl_key_estado + "_x"
            st.rerun()

        if run_clicked and uploaded is not None:
            pbar = st.progress(0, text="Lendo arquivo .txt...")
            try:
                raw_bytes = uploaded.getvalue()
                # Decodificação robusta (UTF-8 -> Latin-1 fallback)
                try:
                    text = raw_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    text = raw_bytes.decode("latin-1")

                pbar.progress(35, text="Convertendo para DataFrame...")
                df = parse_estado_cuenta_txt(text)

                # ======== LINHA TOTAL (mantendo dtype numérico) ========
                if df is not None and not df.empty:
                    numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
                    for c in numeric_cols:
                        df[c] = pd.to_numeric(df[c], errors="coerce")
                    totals = {c: float(np.nansum(df[c].values)) for c in numeric_cols}
                    total_row = {col: "" for col in df.columns}
                    total_row["Descripción"] = "TOTAL"
                    for c in numeric_cols:
                        total_row[c] = totals[c]
                    df = pd.concat([df, pd.DataFrame([total_row], columns=df.columns)], ignore_index=True)

                # ======== VISUAL: formatação com column_config (sem Styler) ========
                numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
                pbar.progress(70, text="Preparando visualização...")

                if df is None or df.empty:
                    st.warning("Nenhuma linha válida encontrada no arquivo.")
                    pbar.progress(0, text="Aguardando...")
                    return

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
                    xlsx_bytes = to_xlsx_bytes_numformat(
                        df, sheet_name="EstadoCuenta", numeric_cols=numeric_cols
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
    # Modo: Plantilla de Gastos (.xlsx/.xls) — SEM totalizador (sem Styler)
    # -------------------------------------------------------------------------
    elif mode == "plantilla":
        upl_key_pg = st.session_state["aag_state"].setdefault("uploader_key_pg", "aag_pg_upl_1")

        st.caption("Carregue o arquivo **Excel** da *Plantilla de Gastos* (primeira aba será lida).")
        uploaded_xl = st.file_uploader(
            "Selecionar arquivo (.xlsx ou .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key=upl_key_pg,
            help="A coluna 'Amount' será formatada como 111,111,111.00 na visualização e no XLSX.",
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", use_container_width=True, disabled=(uploaded_xl is None))
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key_pg"] = upl_key_pg + "_x"
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

                # Garante tipo numérico em Amount
                df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

                # VISUAL: apenas Amount com 2 casas (sem milhar no componente)
                pbar.progress(70, text="Preparando visualização...")
                st.success("Arquivo carregado com sucesso.")
                st.dataframe(
                    df_pg,
                    use_container_width=True,
                    height=550,
                    column_config={
                        str(amount_col): st.column_config.NumberColumn(format="%.2f")
                    },
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
                    xlsx_bytes = to_xlsx_bytes_numformat(
                        df_pg, sheet_name="PlantillaGastos", numeric_cols=[amount_col]
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
        st.caption(
            "Carregue o **Estado de Cuenta (.txt)** e a **Plantilla de Gastos (.xlsx/.xls)**. "
            "Vamos consolidar por conta e comparar os saldos."
        )

        col_u1, col_u2 = st.columns(2)
        with col_u1:
            up_txt = st.file_uploader(
                "Estado de Cuenta (.txt)", type=["txt"],
                key="aag_analise_txt", help="Relatório 'Listado de Saldos' em texto."
            )
        with col_u2:
            up_xl = st.file_uploader(
                "Plantilla de Gastos (.xlsx/.xls)", type=["xlsx", "xls"],
                key="aag_analise_xl", help="Primeira aba será lida."
            )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button(
                "▶️ Executar análise",
                type="primary",
                use_container_width=True,
                disabled=not (up_txt and up_xl),
            )
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            # Reseta apenas os uploaders desta seção
            st.session_state["aag_analise_txt"] = None
            st.session_state["aag_analise_xl"] = None
            st.rerun()

        if run_clicked and up_txt and up_xl:
            def _norm_conta(x) -> str:
                """Normaliza o número da conta: deixa apenas dígitos e remove zeros à esquerda."""
                s = re.sub(r"\D", "", str(x))
                s = s.lstrip("0")
                return s if s else ""

            pbar = st.progress(0, text="Lendo Estado de Cuenta (.txt)...")

            # -----------------------
            # 1) Estado de Cuenta
            # -----------------------
            try:
                raw = up_txt.getvalue()
                try:
                    text = raw.decode("utf-8")
                except UnicodeDecodeError:
                    text = raw.decode("latin-1")

                df_ec = parse_estado_cuenta_txt(text)  # reaproveita seu parser
                if df_ec is None or df_ec.empty:
                    st.error("Estado de Cuenta sem linhas válidas.")
                    st.stop()

                # Normaliza colunas-chave
                if "CTA" not in df_ec.columns or "Período" not in df_ec.columns:
                    st.error("Estado de Cuenta não contém as colunas esperadas: 'CTA' e 'Período'.")
                    st.stop()

                df_ec["CTA"] = df_ec["CTA"].apply(_norm_conta)
                df_ec["Período"] = pd.to_numeric(df_ec["Período"], errors="coerce").fillna(0.0)

                # Remove linhas sem conta (ex.: totalizadores que não têm CTA)
                df_ec = df_ec[df_ec["CTA"].astype(str).str.len() > 0]

                # Agrega por conta (CTA) somando Período
                df_ec_agg = (
                    df_ec.groupby("CTA", as_index=False)["Período"]
                    .sum()
                    .rename(columns={"CTA": "Conta", "Período": "Saldo_Estado"})
                )
                pbar.progress(35, text="Lendo Plantilla de Gastos (.xlsx/.xls)...")

            except Exception as e:
                st.error("Erro ao processar o Estado de Cuenta.")
                st.exception(e)
                st.stop()

            # -----------------------
            # 2) Plantilla de Gastos
            # -----------------------
            try:
                name = getattr(up_xl, "name", "").lower()
                engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
                df_pg = pd.read_excel(up_xl, sheet_name=0, engine=engine)

                # Detecta colunas 'Cuenta' e 'Amount' (case-insensitive)
                def _find_col(df, target):
                    for c in df.columns:
                        if str(c).strip().lower() == target:
                            return c
                    # fallback: contém
                    cand = [c for c in df.columns if target in str(c).strip().lower()]
                    return cand[0] if cand else None

                cuenta_col = _find_col(df_pg, "cuenta")
                amount_col = _find_col(df_pg, "amount")

                if cuenta_col is None or amount_col is None:
                    st.error("Plantilla não contém as colunas esperadas: 'Cuenta' e 'Amount'.")
                    st.stop()

                # Normaliza e soma Amount por Cuenta
                df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce").fillna(0.0)
                df_pg["__conta__"] = df_pg[cuenta_col].apply(_norm_conta)
                df_pg = df_pg[df_pg["__conta__"].astype(str).str.len() > 0]

                df_pg_agg = (
                    df_pg.groupby("__conta__", as_index=False)[amount_col]
                    .sum()
                    .rename(columns={"__conta__": "Conta", amount_col: "Valor_Plantilla"})
                )
                pbar.progress(70, text="Comparando saldos...")

            except Exception as e:
                st.error("Erro ao processar a Plantilla de Gastos.")
                st.exception(e)
                st.stop()

            # -----------------------
            # 3) Comparação
            # -----------------------
            try:
                df_cmp = pd.merge(df_ec_agg, df_pg_agg, on="Conta", how="outer")
                for c in ["Saldo_Estado", "Valor_Plantilla"]:
                    df_cmp[c] = pd.to_numeric(df_cmp[c], errors="coerce").fillna(0.0)

                df_cmp["Diferença"] = (df_cmp["Valor_Plantilla"] - df_cmp["Saldo_Estado"]).round(2)

                # Tolerância (padrão 0,01). Ajuste se quiser.
                tol = 0.01
                df_cmp["Divergente?"] = df_cmp["Diferença"].abs() > tol

                # Ordena por maior diferença absoluta primeiro
                df_cmp = df_cmp.sort_values(df_cmp["Diferença"].abs(), ascending=False).reset_index(drop=True)

                pbar.progress(85, text="Preparando visualização...")

                # Métricas rápidas
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Contas (Estado)", f"{df_ec_agg['Conta'].nunique():,}".replace(",", "."))
                with c2:
                    st.metric(
                        "Soma Estado",
                        f"{df_ec_agg['Saldo_Estado'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    )
                with c3:
                    st.metric(
                        "Soma Plantilla",
                        f"{df_pg_agg['Valor_Plantilla'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    )

                # Filtro: apenas divergentes
                only_div = st.checkbox("Mostrar apenas contas com divergência", value=True)
                df_show = df_cmp[df_cmp["Divergente?"]] if only_div else df_cmp

                # Apresentação
                st.dataframe(
                    df_show[["Conta", "Saldo_Estado", "Valor_Plantilla", "Diferença", "Divergente?"]],
                    use_container_width=True, height=520,
                    column_config={
                        "Saldo_Estado": st.column_config.NumberColumn(format="%.2f"),
                        "Valor_Plantilla": st.column_config.NumberColumn(format="%.2f"),
                        "Diferença": st.column_config.NumberColumn(format="%.2f"),
                    },
                )

                pbar.progress(95, text="Gerando arquivos para download...")

                # Downloads
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
                    # Reaproveita helper para XLSX com máscara numérica
                    xlsx_bytes = to_xlsx_bytes_numformat(
                        df_show[["Conta", "Saldo_Estado", "Valor_Plantilla", "Diferença", "Divergente?"]],
                        sheet_name="Analise",
                        numeric_cols=["Saldo_Estado", "Valor_Plantilla", "Diferença"],
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

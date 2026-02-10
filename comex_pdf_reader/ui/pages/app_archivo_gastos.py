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
        if st.button("Analise", use_container_width=True):
            _set_mode("asientos")

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
            help="A coluna 'Amount' será formatada como 111,111,111.00 na visualização e no XLSX.",
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

                # Garante tipo numérico em Amount
                df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

                # Salva o DF para uso na aba Analise
                st.session_state["aag_plantilla_df"] = df_pg.copy()

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

        # Parâmetros (rótulo atualizado)
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
                .rename(columns={"CTA": "Conta", "Período": "Saldo_Estado_Cuenta"})
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
                .rename(columns={"__conta__": "Conta", amount_col: "Saldo_Plantilla_Gastos"})
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
            df_cmp = pd.merge(df_ec_agg, df_pg_agg, on="Conta", how="outer")
            for c in ["Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos"]:
                df_cmp[c] = pd.to_numeric(df_cmp[c], errors="coerce").fillna(0.0)

            df_cmp["Diferença"] = (df_cmp["Saldo_Plantilla_Gastos"] - df_cmp["Saldo_Estado_Cuenta"]).round(2)

            # Divergente? (usada apenas para filtro, não exibida)
            df_cmp["_div"] = df_cmp["Diferença"].abs() > float(tol)

            # Ordenação por maior diferença (abs)
            df_cmp = df_cmp.sort_values(by="Diferença", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)

            pbar.progress(90, text="Preparando visualização...")

            # Métricas rápidas (rótulos atualizados)
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Contas (Estado)", f"{df_ec_agg['Conta'].nunique():,}".replace(",", "."))
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

            # Não exibir a coluna de controle "_div"
            df_show = df_show[["Conta", "Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos", "Diferença"]]

            # Apresentação
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
                xlsx_bytes = to_xlsx_bytes_numformat(
                    df_show,
                    sheet_name="Analise",
                    numeric_cols=["Saldo_Estado_Cuenta", "Saldo_Plantilla_Gastos", "Diferença"],
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

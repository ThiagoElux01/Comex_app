# =====================================================================
# app_archivo_gastos.py — versão completa atualizada com:
# - Estado de Cuenta
# - Plantilla Gastos
# - Analise
# - Upload de Contas
# - EXTRAÇÃO AUTOMÁTICA DE CUENTA DO CABEÇALHO
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
    if "aag_state" not in st.session_state or not isinstance(st.session_state["aag_state"], dict):
        st.session_state["aag_state"] = {}
    aag = st.session_state["aag_state"]

    aag.setdefault("uploader_key_estado", "aag_estado_upl_1")
    aag.setdefault("uploader_key_pg", "aag_pg_upl_1")
    aag.setdefault("uploader_key_contas", "aag_contas_upl_1")

    aag.setdefault("aag_contas_dfs", {})

    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"

def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode

# ---------------------------------------------------------------------
# Função NOVA — Extrair Cuenta do cabeçalho TXT
# ---------------------------------------------------------------------
def extract_cuenta_from_text(text: str) -> str | None:
    """
    Procura algo como:
    'N° de cta. 121201 -- 121201'
    E retorna apenas o número da conta.
    """
    m = re.search(r"N[°º]?\s*de\s*cta\.?\s*([0-9]{3,})", text, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None

# ---------------------------------------------------------------------
# Parsers do Estado de Cuenta
# ---------------------------------------------------------------------
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

def parse_estado_cuenta_txt(texto: str) -> pd.DataFrame:
    linhas = texto.splitlines()

    # EXTRAIR CUENTA DO CABEÇALHO
    cuenta = extract_cuenta_from_text(texto)

    # Encontrar início da tabela
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
        descr = left[len(cta):].strip()

        sal_ob, saldo_ob, periodo, saldo_cb = (_clean_num(x) for x in m.groups())

        dados.append([cta, descr, sal_ob, saldo_ob, periodo, saldo_cb])

    cols = ["CTA", "Descripción", "Sal OB", "Saldo OB", "Período", "Saldo CB"]
    df = pd.DataFrame(dados, columns=cols)

    # converter numéricos
    for c in ["Sal OB", "Saldo OB", "Período", "Saldo CB"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # -------- NOVA COLUNA --------
    df["Cuenta"] = cuenta if cuenta else ""

    return df

# ---------------------------------------------------------------------
# Export XLSX com máscara numérica
# ---------------------------------------------------------------------
def to_xlsx_bytes_numformat(df: pd.DataFrame, sheet_name: str, numeric_cols: list[str]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        for col_name in numeric_cols:
            if col_name not in df.columns:
                continue
            idx = df.columns.get_loc(col_name) + 1
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"

        BLUE = "FF0077B6"
        fill_blue = PatternFill("solid", start_color=BLUE, end_color=BLUE)
        font_white = Font(color="FFFFFF", bold=True)

        for c in ws[1]:
            c.fill = fill_blue
            c.font = font_white

        for col_idx in range(1, ws.max_column + 1):
            max_len = 10
            for r in range(1, ws.max_row + 1):
                val = ws.cell(row=r, column=col_idx).value
                if val is None:
                    continue
                s = f"{val:,.2f}" if isinstance(val, (int,float)) else str(val)
                max_len = max(max_len, len(s))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------
# Página principal
# ---------------------------------------------------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Archivo Gastos")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("Estado de Cuenta", use_container_width=True):
            _set_mode("estado")
    with col2:
        if st.button("Plantilla Gastos", use_container_width=True):
            _set_mode("plantilla")
    with col3:
        if st.button("Analise", use_container_width=True):
            _set_mode("asientos")
    with col4:
        if st.button("Upload de Contas", use_container_width=True):
            _set_mode("contas")

    mode = st.session_state["aag_mode"]
    st.divider()

    # ================================================================
    # MODO: ESTADO DE CUENTA
    # ================================================================
    if mode == "estado":
        upl_key = st.session_state["aag_state"]["uploader_key_estado"]

        st.caption("Carregue o arquivo .txt de Estado de Cuenta.")
        uploaded = st.file_uploader(
            "Selecionar arquivo",
            type=["txt"],
            accept_multiple_files=False,
            key=upl_key
        )

        col_run, col_clear = st.columns([2,1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", disabled=(uploaded is None))
        with col_clear:
            if st.button("Limpar"):
                st.session_state["aag_state"]["uploader_key_estado"] = upl_key + "_x"
                st.session_state.pop("aag_estado_df", None)
                st.rerun()

        if run_clicked and uploaded:
            pbar = st.progress(0, "Lendo arquivo .txt...")

            raw = uploaded.getvalue()
            try:
                text = raw.decode("utf-8")
            except:
                text = raw.decode("latin-1")

            pbar.progress(40, "Convertendo para DataFrame...")
            df_base = parse_estado_cuenta_txt(text)

            if df_base.empty:
                st.error("Nenhuma linha válida encontrada.")
                return

            st.session_state["aag_estado_df"] = df_base.copy()

            # Adicionar linha TOTAL
            df = df_base.copy()
            numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
            totals = {c: float(np.nansum(df[c])) for c in numeric_cols}

            total_row = {col: "" for col in df.columns}
            total_row["Descripción"] = "TOTAL"
            for c in numeric_cols:
                total_row[c] = totals[c]

            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

            pbar.progress(80, "Mostrando dados...")
            st.dataframe(df, use_container_width=True, height=550)

            pbar.progress(95, "Gerando download...")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "CSV",
                    data=df.to_csv(index=False).encode("utf-8"),
                    file_name="estado_de_cuenta.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with c2:
                xlsx_bytes = to_xlsx_bytes_numformat(df, "EstadoCuenta", numeric_cols)
                st.download_button(
                    "XLSX",
                    data=xlsx_bytes,
                    file_name="estado_de_cuenta.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            pbar.progress(100, "Concluído.")

    # ================================================================
    # MODO: PLANTILLA DE GASTOS
    # ================================================================
    elif mode == "plantilla":
        upl_key_pg = st.session_state["aag_state"]["uploader_key_pg"]

        st.caption("Carregue o arquivo Excel da Plantilla de Gastos.")
        uploaded_xl = st.file_uploader(
            "Selecionar arquivo",
            type=["xlsx","xls"],
            accept_multiple_files=False,
            key=upl_key_pg
        )

        col_run, col_clear = st.columns([2,1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", disabled=(uploaded_xl is None))
        with col_clear:
            if st.button("Limpar"):
                st.session_state["aag_state"]["uploader_key_pg"] = upl_key_pg + "_x"
                st.session_state.pop("aag_plantilla_df", None)
                st.rerun()

        if run_clicked and uploaded_xl:
            pbar = st.progress(0, "Lendo arquivo Excel...")

            name = uploaded_xl.name.lower()
            engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"

            df_pg = pd.read_excel(uploaded_xl, engine=engine)

            amount_col = None
            for c in df_pg.columns:
                if str(c).strip().lower() == "amount":
                    amount_col = c
                    break
            if amount_col is None:
                candidates = [c for c in df_pg.columns if "amount" in str(c).lower()]
                if candidates:
                    amount_col = candidates[0]

            if amount_col is None:
                st.error("Coluna 'Amount' não encontrada.")
                return

            df_pg[amount_col] = pd.to_numeric(df_pg[amount_col], errors="coerce")

            st.session_state["aag_plantilla_df"] = df_pg.copy()

            pbar.progress(80, "Mostrando...")
            st.dataframe(df_pg, use_container_width=True, height=550)

            pbar.progress(95, "Gerando downloads...")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "CSV",
                    data=df_pg.to_csv(index=False).encode("utf-8"),
                    file_name="plantilla_gastos.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with c2:
                xlsx_bytes = to_xlsx_bytes_numformat(df_pg, "PlantillaGastos", [amount_col])
                st.download_button(
                    "XLSX",
                    data=xlsx_bytes,
                    file_name="plantilla_gastos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            pbar.progress(100, "Concluído.")

    # ================================================================
    # MODO: ANALISE
    # ================================================================
    elif mode == "asientos":
        st.subheader("🔍 Analise: Estado de Cuenta x Plantilla")

        df_ec = st.session_state.get("aag_estado_df", None)
        df_pg = st.session_state.get("aag_plantilla_df", None)

        faltando = []
        if df_ec is None:
            faltando.append("Estado de Cuenta")
        if df_pg is None:
            faltando.append("Plantilla de Gastos")

        if faltando:
            st.warning("Para analisar, carregue: " + ", ".join(faltando))
            return

        tol = st.number_input("Valor de Tolerância", min_value=0.00, value=0.01, step=0.01)

        def _norm_conta(x):
            s = re.sub(r"\D", "", str(x))
            return s.lstrip("0") or ""

        pbar = st.progress(0, "Processando Estado de Cuenta...")

        # Estado
        df_ec_proc = df_ec.copy()
        df_ec_proc["CTA"] = df_ec_proc["CTA"].apply(_norm_conta)
        df_ec_proc["Período"] = pd.to_numeric(df_ec_proc["Período"], errors="coerce").fillna(0.0)
        df_ec_proc = df_ec_proc[df_ec_proc["CTA"].str.len() > 0]

        df_ec_agg = (
            df_ec_proc.groupby("CTA", as_index=False)["Período"]
            .sum()
            .rename(columns={"CTA":"Cuenta","Período":"Saldo_Estado_Cuenta"})
        )

        pbar.progress(40, "Processando Plantilla...")

        # Plantilla
        def _find_col(df, target):
            for c in df.columns:
                if str(c).strip().lower() == target:
                    return c
            cand = [c for c in df.columns if target in str(c).lower()]
            return cand[0] if cand else None

        cuenta_col = _find_col(df_pg, "cuenta")
        amount_col = _find_col(df_pg, "amount")

        if cuenta_col is None or amount_col is None:
            st.error("Plantilla não contém 'Cuenta' e/ou 'Amount'.")
            return

        df_pg_proc = df_pg.copy()
        df_pg_proc[amount_col] = pd.to_numeric(df_pg_proc[amount_col], errors="coerce").fillna(0.0)
        df_pg_proc["__conta__"] = df_pg_proc[cuenta_col].apply(_norm_conta)
        df_pg_proc = df_pg_proc[df_pg_proc["__conta__"].str.len() > 0]

        df_pg_agg = (
            df_pg_proc.groupby("__conta__", as_index=False)[amount_col]
            .sum()
            .rename(columns={"__conta__":"Cuenta", amount_col:"Saldo_Plantilla_Gastos"})
        )

        pbar.progress(70, "Comparando...")

        # Merge
        df_cmp = pd.merge(df_ec_agg, df_pg_agg, on="Cuenta", how="outer")
        for c in ["Saldo_Estado_Cuenta","Saldo_Plantilla_Gastos"]:
            df_cmp[c] = pd.to_numeric(df_cmp[c], errors="coerce").fillna(0.0)

        df_cmp["Diferença"] = (df_cmp["Saldo_Plantilla_Gastos"] - df_cmp["Saldo_Estado_Cuenta"]).round(2)
        df_cmp["_div"] = df_cmp["Diferença"].abs() > float(tol)

        df_cmp["_cuenta_num"] = pd.to_numeric(df_cmp["Cuenta"], errors="coerce")
        df_cmp = df_cmp.sort_values("_cuenta_num").drop(columns=["_cuenta_num"])

        pbar.progress(90, "Exibindo...")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Contas (Estado)", df_ec_agg["Cuenta"].nunique())
        with c2:
            st.metric("Soma Estado", df_ec_agg["Saldo_Estado_Cuenta"].sum())
        with c3:
            st.metric("Soma Plantilla", df_pg_agg["Saldo_Plantilla_Gastos"].sum())

        only_div = st.checkbox("Mostrar apenas divergentes", value=True)
        df_show = df_cmp[df_cmp["_div"]] if only_div else df_cmp

        df_show = df_show[["Cuenta","Saldo_Estado_Cuenta","Saldo_Plantilla_Gastos","Diferença"]]

        st.dataframe(df_show, use_container_width=True, height=550)

        pbar.progress(95, "Gerando downloads...")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "CSV",
                data=df_show.to_csv(index=False).encode("utf-8"),
                file_name="analise_contas.csv",
                mime="text/csv",
                use_container_width=True
            )
        with c2:
            xlsx_bytes = to_xlsx_bytes_numformat(
                df_show, "Analise",
                ["Saldo_Estado_Cuenta","Saldo_Plantilla_Gastos","Diferença"]
            )
            st.download_button(
                "XLSX",
                data=xlsx_bytes,
                file_name="analise_contas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        pbar.progress(100, "Concluído.")

    # ================================================================
    # MODO: UPLOAD DE CONTAS
    # ================================================================
    elif mode == "contas":
        st.subheader("📂 Upload de Contas — Múltiplos Formatos")

        upl_key = st.session_state["aag_state"]["uploader_key_contas"]

        uploaded_files = st.file_uploader(
            "Selecione arquivos",
            type=["txt","csv","xlsx","xls"],
            accept_multiple_files=True,
            key=upl_key
        )

        col_run, col_clear = st.columns([2,1])
        with col_run:
            run_clicked = st.button("▶️ Processar", type="primary",
                disabled=(not uploaded_files))
        with col_clear:
            if st.button("Limpar"):
                st.session_state["aag_state"]["uploader_key_contas"] = upl_key + "_x"
                st.session_state["aag_state"]["aag_contas_dfs"] = {}
                st.rerun()

        if run_clicked and uploaded_files:
            pbar = st.progress(0, "Processando arquivos...")
            dfs = {}
            total = len(uploaded_files)

            for i, file in enumerate(uploaded_files):
                fname = file.name.lower()
                try:
                    if fname.endswith(".txt"):
                        raw = file.getvalue()
                        try:
                            text = raw.decode("utf-8")
                        except:
                            text = raw.decode("latin-1")
                        df = pd.DataFrame({"linha": text.splitlines()})

                    elif fname.endswith(".csv"):
                        df = pd.read_csv(file)

                    elif fname.endswith(".xlsx"):
                        df = pd.read_excel(file, engine="openpyxl")

                    elif fname.endswith(".xls"):
                        df = pd.read_excel(file, engine="xlrd")

                    else:
                        st.warning(f"Formato não reconhecido: {fname}")
                        continue

                    dfs[file.name] = df

                except Exception as e:
                    st.error(f"Erro ao processar {file.name}")
                    st.exception(e)

                pbar.progress(int((i+1)/total * 100))

            st.session_state["aag_state"]["aag_contas_dfs"] = dfs
            st.success("Arquivos processados com sucesso!")
            pbar.progress(100, "Concluído.")

        # Exibir
        dfs = st.session_state["aag_state"]["aag_contas_dfs"]
        if dfs:
            st.write("### 📄 Arquivos carregados")
            for fname, df in dfs.items():
                st.write(f"#### 📌 {fname}")
                st.dataframe(df, use_container_width=True, height=350)


# ---------------------------------------------------------------------
# Executar
# ---------------------------------------------------------------------
if __name__ == "__main__":
    render()

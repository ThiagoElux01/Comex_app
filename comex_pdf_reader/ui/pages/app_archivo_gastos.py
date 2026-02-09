# ui/pages/app_archivo_gastos.py
import re
from io import BytesIO
import streamlit as st
import pandas as pd

# Reaproveita helper de exportação XLSX da Aplicación Comex
# (definido em ui/pages/process_pdfs.py)
from ui.pages.process_pdfs import to_xlsx_bytes  # mesmo padrão de exportação (autofit/estilo)

# ------------------------------------------------------------
# Estado e helpers
# ------------------------------------------------------------
def _ensure_state():
    if "aag_state" not in st.session_state:
        st.session_state["aag_state"] = {
            "uploader_key": "aag_uploader_1",
            "last_action": None,      # "estado" | "plantilla" | "asientos"
        }
    if "aag_mode" not in st.session_state:
        st.session_state["aag_mode"] = "estado"  # default na primeira carga

def _set_mode(mode: str):
    st.session_state["aag_mode"] = mode

# ------------------------------------------------------------
# Parsers
# ------------------------------------------------------------
_NUM = r"(-?\d[\d,]*\.\d{2}-?)"   # número com milhares e 2 decimais; pode terminar com '-' (negativo)

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
        if "CTA" in ln and "Descripción" in ln:
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
        left = raw[:m.start()].rstrip()
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

    # Tipos numéricos garantidos (caso algo tenha escapado)
    for c in ["Sal OB", "Saldo OB", "Período", "Saldo CB"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df

# ------------------------------------------------------------
# Página
# ------------------------------------------------------------
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
        if st.button("Asientos", use_container_width=True):
            _set_mode("asientos")

    mode = st.session_state["aag_mode"]
    st.divider()

    # --------------------------------------------------------
    # Modo: Estado de Cuenta  (ativo agora)
    # --------------------------------------------------------
    if mode == "estado":
        st.caption("Carregue o arquivo **.txt** de *Listado de Saldos* para visualização e export.")
        uploaded = st.file_uploader(
            "Selecionar arquivo (.txt)",
            type=["txt"],
            accept_multiple_files=False,
            key=st.session_state["aag_state"]["uploader_key"],
            help="Ex.: relatório 'Listado de Saldos' exportado do sistema."
        )

        col_run, col_clear = st.columns([2, 1])
        with col_run:
            run_clicked = st.button("▶️ Executar", type="primary", use_container_width=True, disabled=(uploaded is None))
        with col_clear:
            clear_clicked = st.button("Limpar", use_container_width=True)

        if clear_clicked:
            st.session_state["aag_state"]["uploader_key"] = st.session_state["aag_state"]["uploader_key"] + "_x"
            st.rerun()

        if run_clicked and uploaded is not None:
            pbar = st.progress(0, text="Lendo arquivo .txt...")
            try:
                raw_bytes = uploaded.getvalue()
                # Decodificação robusta (primeiro UTF-8, se falhar cai para Latin-1)
                try:
                    text = raw_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    text = raw_bytes.decode("latin-1")

                pbar.progress(35, text="Convertendo para DataFrame...")
                df = parse_estado_cuenta_txt(text)

                # ======== AJUSTE: adiciona linha de totais no final ========
                if df is not None and not df.empty:
                    # 1) Garante que as colunas numéricas são float
                    numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
                    for c in numeric_cols:
                        df[c] = pd.to_numeric(df[c], errors="coerce")

                    # 2) (Opcional) adiciona uma linha em branco antes do TOTAL (visual)
                    blank_row = {col: "" for col in df.columns}
                    df = pd.concat([df, pd.DataFrame([blank_row], columns=df.columns)], ignore_index=True)

                    # 3) Calcula o TOTAL
                    totals_row = {col: "" for col in df.columns}
                    totals_row["Descripción"] = "TOTAL"
                    for c in numeric_cols:
                        totals_row[c] = float(df[c].sum(skipna=True))

                    # 4) Concatena o TOTAL e arredonda
                    df = pd.concat([df, pd.DataFrame([totals_row], columns=df.columns)], ignore_index=True)
                    df[numeric_cols] = df[numeric_cols].round(2)
                # ======== FIM DO AJUSTE ========

                pbar.progress(70, text="Preparando visualização...")
                if df is None or df.empty:
                    st.warning("Nenhuma linha válida encontrada no arquivo.")
                    pbar.progress(0, text="Aguardando...")
                    return

                st.success("Arquivo processado com sucesso.")
                st.dataframe(df, use_container_width=True, height=550)

                pbar.progress(90, text="Gerando arquivos para download...")
                # Downloads
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
                    xlsx_bytes = to_xlsx_bytes(df, sheet_name="EstadoCuenta")
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

    # --------------------------------------------------------
    # Placeholders para os demais (prontos para receber lógica)
    # --------------------------------------------------------
    elif mode == "plantilla":
        st.info("🧩 *Plantilla Gastos* — em breve conectaremos a lógica aqui.")
    elif mode == "asientos":
        st.info("📒 *Asientos* — em breve conectaremos a lógica aqui.")

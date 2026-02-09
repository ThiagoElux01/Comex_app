# ui/pages/app_archivo_gastos.py
import re
import numpy as np
import streamlit as st
import pandas as pd

from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers, PatternFill, Font

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
# Export XLSX com máscara numérica #,##0.00 (mantém tipo numérico)
# ------------------------------------------------------------
def to_xlsx_bytes_numformat(df: pd.DataFrame, sheet_name: str, numeric_cols: list[str]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]

        # Aplica máscara #,##0.00 nas colunas numéricas
        for col_name in numeric_cols:
            if col_name not in df.columns:
                continue
            col_idx = df.columns.get_loc(col_name) + 1  # 1-based
            for row in range(2, ws.max_row + 1):  # pulando cabeçalho
                cell = ws.cell(row=row, column=col_idx)
                # aplica formato só se for número
                if isinstance(cell.value, (int, float)) and cell.value is not None:
                    cell.number_format = '#,##0.00'

        # Pintar cabeçalho (igual vibe da Aplicación Comex)
        BLUE = "FF0077B6"
        WHITE = "FFFFFFFF"
        fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
        font_white_bold = Font(color=WHITE, bold=True)
        for cell in ws[1]:
            cell.fill = fill_blue
            cell.font = font_white_bold

        # Ajuste de largura simples
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
    # Modo: Estado de Cuenta
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

                # ======== LINHA TOTAL (mantendo dtype numérico) ========
                if df is not None and not df.empty:
                    numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]

                    # 1) Garante float nas colunas numéricas
                    for c in numeric_cols:
                        df[c] = pd.to_numeric(df[c], errors="coerce")

                    # 2) Soma com numpy (ignora NaN)
                    totals = {c: float(np.nansum(df[c].values)) for c in numeric_cols}

                    # 3) Cria linha TOTAL
                    total_row = {col: "" for col in df.columns}
                    total_row["Descripción"] = "TOTAL"
                    for c in numeric_cols:
                        total_row[c] = totals[c]

                    # 4) Concatena TOTAL
                    df = pd.concat([df, pd.DataFrame([total_row], columns=df.columns)], ignore_index=True)

                # ======== VISUAL: formatação com Styler (alinha à direita e mostra 111,111,111.00) ========
                numeric_cols = ["Sal OB", "Saldo OB", "Período", "Saldo CB"]
                fmt_dict = {c: "{:,.2f}".format for c in numeric_cols}

                # NB: st.dataframe aceita Styler e preserva alinhamento à direita de números
                styler = (
                    df.style
                    .format(fmt_dict, na_rep="")
                )

                pbar.progress(70, text="Preparando visualização...")
                if df is None or df.empty:
                    st.warning("Nenhuma linha válida encontrada no arquivo.")
                    pbar.progress(0, text="Aguardando...")
                    return

                st.success("Arquivo processado com sucesso.")
                st.dataframe(styler, use_container_width=True, height=550)

                pbar.progress(90, text="Gerando arquivos para download...")
                # Downloads:
                # 1) CSV — mantém numéricos como numéricos (sem formatação de milhar no arquivo)
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        label="Baixar CSV (Estado de Cuenta)",
                        data=df.to_csv(index=False).encode("utf-8"),
                        file_name="estado_de_cuenta.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

                # 2) XLSX — mantém numérico e aplica máscara #,##0.00
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes_numformat(df, sheet_name="EstadoCuenta", numeric_cols=numeric_cols)
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

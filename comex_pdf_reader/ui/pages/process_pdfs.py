# ui/pages/process_pdfs.py
import streamlit as st
from services import pdf_service
from services.tasa_service import atualizar_dataframe_tasa
from ui.pages import downloads_page

# --- NOVO: Import protegido do Adicionales ---
ADICIONALES_AVAILABLE = True
ADICIONALES_ERR = None
try:
    from services.adicionales_service import process_adicionales_streamlit
except Exception as e:
    ADICIONALES_AVAILABLE = False
    ADICIONALES_ERR = e

if not ADICIONALES_AVAILABLE:
    st.warning(
        "Módulo **Gastos Adicionales** não pôde ser carregado. "
        "Verifique `services/adicionales_service.py` e dependências (ex.: `PyMuPDF`)."
    )
    with st.expander("Detalhes técnicos do erro (Adicionales)"):
        st.exception(ADICIONALES_ERR)

# -----------------------------
# Imports protegidos (diagnóstico no app)
# -----------------------------
# Import protegido do Percepciones
PERC_AVAILABLE = True
PERC_ERR = None
try:
    from services.percepcion_service import process_percepcion_streamlit
except Exception as e:
    PERC_AVAILABLE = False
    PERC_ERR = e

if not PERC_AVAILABLE:
    st.warning(
        "Módulo **Percepciones** não pôde ser carregado. "
        "Verifique `services/percepcion_service.py` e dependências (ex.: PyMuPDF)."
    )
    with st.expander("Detalhes técnicos do erro (Percepciones)"):
        st.exception(PERC_ERR)

# Import protegido do DUAS (para não “matar” o tab se faltar dependência)
DUAS_AVAILABLE = True
DUAS_ERR = None
try:
    from services.duas_service import process_duas_streamlit
except Exception as e:
    DUAS_AVAILABLE = False
    DUAS_ERR = e

# AVISO se o módulo DUAS não carregou (mas mantém os botões visíveis)
if not DUAS_AVAILABLE:
    st.warning(
        "O módulo **DUAS** não pôde ser carregado. "
        "Verifique `services/duas_service.py` e dependências (ex.: `pdfplumber`)."
    )
    with st.expander("Detalhes técnicos do erro (DUAS)"):
        st.exception(DUAS_ERR)  # <- mostra o stack-trace real

# --- NOVO: Import protegido do Externos (segue o mesmo padrão dos demais) ---
EXTERNOS_AVAILABLE = True
EXTERNOS_ERR = None
try:
    from services.externos_service import process_externos_streamlit
except Exception as e:
    EXTERNOS_AVAILABLE = False
    EXTERNOS_ERR = e

if not EXTERNOS_AVAILABLE:
    st.warning(
        "Módulo **Externos** não pôde ser carregado. "
        "Verifique `services/externos_service.py` e dependências (ex.: `PyMuPDF`)."
    )
    with st.expander("Detalhes técnicos do erro (Externos)"):
        st.exception(EXTERNOS_ERR)

# -----------------------------
# Utilidades
# -----------------------------
from io import BytesIO
import pandas as pd

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# INCLUSÃO: utilitário para ajustar largura ("autofit") das colunas no Excel
from openpyxl.utils import get_column_letter

def _autofit_worksheet(ws, font_padding: float = 1.2, min_width: float = 8.0, max_width: float = 60.0):
    """
    Ajusta a largura das colunas de uma planilha openpyxl com base no maior texto
    (entre cabeçalho e células). Não existe AutoFit nativo no openpyxl; esta é uma estimativa.
    """
    if ws.max_column is None or ws.max_row is None:
        return
    for col_idx, col in enumerate(
        ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column),
        start=1
    ):
        header = col[0].value if col and col[0] is not None else ""
        max_len = len(str(header)) if header is not None else 0
        for cell in col[1:]:  # ignora o cabeçalho já contabilizado
            val = cell.value
            if val is None:
                continue
            # Representação razoável para floats (evita notação científica gigante)
            text = f"{val:.6g}" if isinstance(val, float) else str(val)
            max_len = max(max_len, len(text))
        width = min(max(max_len * font_padding, min_width), max_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

from openpyxl.styles import PatternFill, Font

def header_paint(ws):
    """
    Pinta o cabeçalho (linha 1) apenas quando o texto for
    exatamente igual (case-sensitive) a um dos nomes definidos.
    """
    BLUE = "FF0077B6"  # ARGB (FF = opacidade total)
    WHITE = "FFFFFFFF"
    fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
    font_white_bold = Font(color=WHITE, bold=True)

    # Lista EXATA (case-sensitive). Só estes serão pintados.
    exact_headers = {
        "source_file",
        "Proveedor",
        "Fornecedor",
        "Proveedor Iscala",
        "Factura",
        "Tipo Doc",
        "Cód. de Autorización",
        "Tipo de Factura",
        "Fecha de Emisión",
        "Moneda",
        "Cod. Moneda",
        "Amount",
        "Tasa",
        "Cuenta",
        "Error",
        "R.U.C",
        "Op. Gravada",
        "COD PROVEEDOR",
        "Declaracion",
        "Fecha",
        "Ad_Valorem",
        "Imp_Prom_Municipal",
        "Imp_Gene_a_las_Ventas",
        "IGV",
        "Percepcion",
        "PEC",
        "COD Moneda",
        "Source_File",
        "No_Liquidacion",
        "CDA",
        "Monto",
        "COD MONEDA",
        "Lineaabajo",
    }

    # Percorre somente a linha 1 (cabeçalho)
    for cell in ws[1]:
        if cell.value is None:
            continue
        header_text = str(cell.value).strip()  # sem lower(), comparação exata
        if header_text in exact_headers:
            cell.fill = fill_blue
            cell.font = font_white_bold

def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Tasa") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        _autofit_worksheet(ws)
        header_paint(ws)  # <- aqui
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# NOVO: criar DF com linhas em branco entre cada registro
# -----------------------------------------------------------------------------
def df_with_blank_spacers(df: pd.DataFrame, blank_rows: int = 3) -> pd.DataFrame:
    """
    Retorna um novo DataFrame onde após cada linha há `blank_rows` linhas vazias.
    Mantém as mesmas colunas; as linhas vazias são None.
    """
    if df is None or df.empty:
        return df.copy()
    blocks = []
    blank = pd.DataFrame([[None] * len(df.columns)], columns=df.columns)
    for _, row in df.iterrows():
        blocks.append(pd.DataFrame([row.values], columns=df.columns))
        for _ in range(blank_rows):
            blocks.append(blank.copy())
    out = pd.concat(blocks, ignore_index=True)
    return out

# -----------------------------------------------------------------------------
# NOVO: gerar XLSX de Externos com duas abas (normal + espaçado)
# -----------------------------------------------------------------------------
def to_xlsx_bytes_externos_duas_abas(
    df_normal: pd.DataFrame,
    sheet_normal: str = "Externos",
    sheet_spaced: str = "Externos_Espacado",
    blank_rows: int = 3
) -> bytes:
    """
    Cria um XLSX com:
    - Aba 1: dados originais de 'Externos'
    - Aba 2: os mesmos dados, mas com 3 linhas em branco abaixo de cada registro
    Aplica auto-ajuste e pintura do cabeçalho em ambas as abas.
    """
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Aba normal
        df_normal.to_excel(writer, index=False, sheet_name=sheet_normal)
        ws1 = writer.book[sheet_normal]
        _autofit_worksheet(ws1)
        header_paint(ws1)

        # Aba espaçada
        df_spaced = df_with_blank_spacers(df_normal, blank_rows=blank_rows)
        df_spaced.to_excel(writer, index=False, sheet_name=sheet_spaced)
        ws2 = writer.book[sheet_spaced]
        _autofit_worksheet(ws2)
        header_paint(ws2)
    buffer.seek(0)
    return buffer.getvalue()

ACTIONS = {
    "externos": "Externos",
    "gastos": "Gastos Adicionales",
    "duas": "Duas",
    "percepciones": "Percepciones",
}

def _ensure_state():
    if "acao_selecionada" not in st.session_state:
        st.session_state.acao_selecionada = None
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = "uploader_none"
    if "tasa_df" not in st.session_state:
        st.session_state.tasa_df = None

def _select_action(action_key: str):
    st.session_state.acao_selecionada = action_key
    st.session_state.uploader_key = f"uploader_{action_key}"

# ================== NOVOS HELPERS (PRN) ==================
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

_num_token = re.compile(r'^\s*[-+]?\d+(?:\.\d+)?\s*$')
_invis = ("\u200b" "\u2060")  # zero-width space / word joiner

def _to_str(x):
    """
    Converte valor para string:
    - None/NaN -> ""
    - numérico -> 2 casas decimais com ponto (ROUND_HALF_UP)
    - demais -> str(x) preservando
    Sanitiza NBSP e caracteres invisíveis comuns.
    """
    if x is None:
        return ""
    s = str(x)
    # Remove NBSP e invisíveis
    s = s.replace("\u00a0", " ")
    for ch in _invis:
        s = s.replace(ch, "")
    s = s.strip()

    if s.lower() in {"nan", "none"}:
        return ""

    if _num_token.match(s):
        try:
            q = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            return f"{q:.2f}"  # sempre 2 casas
        except (InvalidOperation, ValueError):
            return s
    return s

def _fixed_width_line(values, widths):
    """
    Monta uma linha em largura fixa, formatando números com 2 casas
    e preenchendo com espaços à direita.
    """
    out = []
    for v, w in zip(values, widths):
        s = _to_str(v)
        s = s[:w]
        out.append(s + (" " * max(0, w - len(s))))
    return "".join(out)

def _df_to_prn_bytes(rows_values, widths, encoding="cp1252"):
    """
    Converte uma lista de 'values por linha' num PRN de largura fixa.
    Retorna bytes (para usar no st.download_button).
    """
    lines = [_fixed_width_line(vals, widths) for vals in rows_values]
    text = "\r\n".join(lines) + "\r\n"  # CRLF, como o Excel
    return text.encode(encoding, errors="replace")

# ------------------- EXTERNOS: 1ª ABA (equivalente Carga_Financeira) -------------------
def gerar_externos_prn_primeira_aba(xls_file):
    """
    Lê a 1ª aba do arquivo Excel e produz 'Externos.prn' replicando a macro Carga_Financeira:
    - pega linhas 3..1500 pulando de 4 em 4
    - se coluna C (3) não vazia, copia colunas C..Z (3..26) -> 24 colunas
    - aplica larguras A..X conforme a macro
    """
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"

    # dtype=str para preservar o que vier como texto; sanitizamos no get_cell
    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        row_idx = r - 2  # Excel r=2 -> df.iloc[0]
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df.index) and 0 <= col_idx < len(df.columns):
                return _to_str(df.iloc[row_idx, col_idx])  # força sanitização aqui
        except Exception:
            return ""
        return ""

    widths = [10, 25, 6, 6, 6, 16, 16, 2, 5, 16, 3, 2, 30, 6, 3, 3, 8, 3, 6, 4, 16, 16, 3, 6]

    rows_values = []
    for r in range(3, 1501, 4):  # 3 até 1500 pulando de 4
        val_c = get_cell(r, 3)  # já sanitizado
        if val_c != "":
            row_vals = [get_cell(r, c) for c in range(3, 27)]  # C..Z -> 24 colunas
            rows_values.append(row_vals)

    prn_bytes = _df_to_prn_bytes(rows_values, widths, encoding="cp1252")
    return prn_bytes  # para "Externos.prn"

# ------------------- EXTERNOS: 2ª ABA (equivalente Carga_Contabil) -------------------
def gerar_externos_prn_segunda_aba(xls_file):
    """
    Lê a 2ª aba do arquivo Excel e produz 'aexternos.prn' replicando Carga_Contabil:
    - acha linhaLimite: primeira linha, a partir de B2, onde B é erro (#N/A). No pandas, tratamos NaN como quebra.
      -> linhaLimite = (linhaErro - 4); se não achar, usa 1496
    - copia intervalo B2:N(linhaLimite) -> 13 colunas
    - larguras: A..M = [6,3,3,8,3,16,16,2,30,6,15,20,5]
    - regras:
      D == 0 -> limpar (ficar "")
      F vazio ou 0 -> remover linha
    """
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df2 = pd.read_excel(xls_file, sheet_name=1, header=0, dtype=str, engine=engine)

    def get_cell2(r, c):
        row_idx = r - 2
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df2.index) and 0 <= col_idx < len(df2.columns):
                return _to_str(df2.iloc[row_idx, col_idx])  # sanitizado
        except Exception:
            return ""
        return ""

    # 1) Encontrar linhaLimite
    linha_limite = 0
    for r in range(2, 46001):  # B2..B46000
        val_b = get_cell2(r, 2)  # coluna B (já sanitizado)
        if (val_b is None) or (str(val_b).strip() in {"#N/A", "#N/D"}) or (pd.isna(val_b)):
            linha_limite = r - 4
            break
    if linha_limite <= 0:
        linha_limite = 1496

    # 2) Copiar B2:N(linha_limite)
    rows_raw = []
    for r in range(2, max(2, linha_limite) + 1):
        row_vals = [get_cell2(r, c) for c in range(2, 15)]  # 2..14 (B..N) => 13 colunas
        rows_raw.append(row_vals)

    # 3) Regras sobre as colunas:
    # D = idx 3 (0-based), F = idx 5 (0-based) no array rows_raw
    rows_clean = []
    for vals in rows_raw:
        d_val = _to_str(vals[3])
        if d_val.strip() in {"0", "0.0", "0.00"}:
            vals[3] = ""
        f_val = _to_str(vals[5]).strip()
        if f_val in {"", "0", "0.0", "0.00"}:
            continue
        rows_clean.append(vals)

    widths2 = [6, 3, 3, 8, 3, 16, 16, 2, 30, 6, 15, 20, 5]
    prn_bytes = _df_to_prn_bytes(rows_clean, widths2, encoding="cp1252")
    return prn_bytes  # para "aexternos.prn"

# ------------------- ADICIONALES: 1ª ABA (igual Carga_Financeira) -------------------
def gerar_adicionales_prn_primeira_aba(xls_file):
    """
    Lê a 1ª aba do arquivo Excel e produz 'Adicionales.prn':
    - linhas 3..1500 pulando de 4 em 4
    - se coluna C (3) não vazia, pega colunas C..Z (3..26) => 24 colunas
    - aplica larguras A..X idênticas às da macro
    """
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"

    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        row_idx = r - 2  # Excel r=2 -> df.iloc[0]
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df.index) and 0 <= col_idx < len(df.columns):
                return _to_str(df.iloc[row_idx, col_idx])
        except Exception:
            return ""
        return ""

    widths = [10, 25, 6, 6, 6, 16, 16, 2, 5, 16, 3, 2, 30, 6, 3, 3, 8, 3, 6, 4, 16, 16, 3, 6]

    rows_values = []
    for r in range(3, 1501, 4):
        val_c = get_cell(r, 3)  # coluna C
        if val_c != "":
            row_vals = [get_cell(r, c) for c in range(3, 27)]  # C..Z
            rows_values.append(row_vals)

    prn_bytes = _df_to_prn_bytes(rows_values, widths, encoding="cp1252")
    return prn_bytes  # "Adicionales.prn"


# ------------------- ADICIONALES: 1ª ABA (PRNs individuais -> ZIP) -------------------
def gerar_adicionales_zip_primeira_aba(xls_file, zip_name="Adicionales_PRNs.zip"):
    """
    Gera múltiplos PRNs (um por linha) da 1ª aba e retorna um ZIP com todos.
    Nome de cada PRN: <valor_coluna_C_sanitizado>_<seq>.prn
    """
    from zipfile import ZipFile, ZIP_DEFLATED
    from io import BytesIO

    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"

    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        row_idx = r - 2
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df.index) and 0 <= col_idx < len(df.columns):
                return _to_str(df.iloc[row_idx, col_idx])
        except Exception:
            return ""
        return ""

    widths = [10, 25, 6, 6, 6, 16, 16, 2, 5, 16, 3, 2, 30, 6, 3, 3, 8, 3, 6, 4, 16, 16, 3, 6]

    buffer_zip = BytesIO()
    with ZipFile(buffer_zip, mode="w", compression=ZIP_DEFLATED) as zf:
        seq = 1
        for r in range(3, 1501, 4):
            val_c = get_cell(r, 3)  # coluna C
            if val_c == "":
                continue

            row_vals = [get_cell(r, c) for c in range(3, 27)]
            prn_bytes = _df_to_prn_bytes([row_vals], widths, encoding="cp1252")
            # Sanitize para nome de arquivo
            safe_prefix = (val_c or "linha").replace("\\", "_").replace("/", "_").replace(" ", "")
            filename = f"{safe_prefix}_{seq}.prn"
            zf.writestr(filename, prn_bytes)
            seq += 1

    buffer_zip.seek(0)
    return buffer_zip.getvalue(), zip_name  # ZIP em bytes + nome padrão


# ------------------- ADICIONALES: 2ª ABA (igual Aexternos.prn) -------------------
def gerar_adicionales_prn_segunda_aba(xls_file):
    """
    Lê a 2ª aba do Excel e produz 'AAdicionales.prn', reutilizando a mesma regra do Aexternos.prn:
    - encontra linha limite a partir de B2
    - usa intervalo B2:N(linhaLimite) (13 colunas) com filtros (D->limpar 0; remove F vazio/0)
    - larguras A..M = [6,3,3,8,3,16,16,2,30,6,15,20,5]
    """
    return gerar_externos_prn_segunda_aba(xls_file)
# ================== FIM HELPERS (PRN) ==================

# -----------------------------
# Página
# -----------------------------
def render():
    _ensure_state()
    st.subheader("Aplicación Comex")

    tab4, tab2, tab3, tab1, tab5 = st.tabs([
        "📦 Arquivos modelo",
        "🌐 Tasa SUNAT",
        "📁 Arquivo Sharepoint",
        "📥 Processamento local",
        "📝 Transformar .prn"
    ])

    # -------------------------
    # 📥 Processamento local
    # -------------------------
    with tab1:
        st.markdown("#### Ações rápidas")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Externos", key="act_externos", use_container_width=True):
                _select_action("externos")
            if st.button("Gastos Adicionales", key="act_gastos", use_container_width=True):
                _select_action("gastos")
        with col2:
            if st.button("Duas", key="act_duas", use_container_width=True):
                _select_action("duas")
            if st.button("Percepciones", key="act_perc", use_container_width=True):
                _select_action("percepciones")

        if not DUAS_AVAILABLE:
            st.warning(
                "O módulo **DUAS** não pôde ser carregado. "
                "Verifique `services/duas_service.py` e dependências (ex.: `pdfplumber`)."
            )

        has_action = st.session_state.acao_selecionada is not None
        if has_action:
            nome_acao = ACTIONS[st.session_state.acao_selecionada]
            st.info(f"🔧 Fluxo **{nome_acao}** selecionado.")
        else:
            st.caption("Selecione uma ação acima para enviar PDFs e executar o fluxo correspondente.")

        st.divider()

        if has_action:
            uploaded_files = st.file_uploader(
                f"Envie um ou mais arquivos PDF para **{ACTIONS[st.session_state.acao_selecionada]}**",
                type=["pdf"],
                accept_multiple_files=True,
                key=st.session_state.uploader_key,
                help="Os arquivos enviados serão processados pelo fluxo selecionado."
            )
            col_run, col_clear = st.columns([2, 1])
            with col_run:
                run_clicked = st.button(
                    "▶️ Executar",
                    key="action_run",
                    type="primary",
                    use_container_width=True,
                    disabled=not uploaded_files
                )
            with col_clear:
                clear_clicked = st.button("Limpar seleção", key="action_clear", use_container_width=True)

            if clear_clicked:
                st.session_state.acao_selecionada = None
                st.session_state.uploader_key = "uploader_none"
                st.rerun()

            if run_clicked and uploaded_files:
                acao = st.session_state.acao_selecionada
                nome_acao = ACTIONS[acao]
                status = st.empty()
                progress = st.progress(0, text=f"Iniciando fluxo {nome_acao}...")

                if acao == "duas":
                    cambio_df = st.session_state.get("tasa_df")
                    if cambio_df is None or getattr(cambio_df, "empty", True):
                        st.warning("Para calcular **Tasa**, primeiro atualize no tab **🌐 Tasa SUNAT**. O processamento seguirá sem Tasa.")
                    if not DUAS_AVAILABLE:
                        st.error("DUAS indisponível: confira dependências e arquivo `services/duas_service.py`.")
                    else:
                        df_final = process_duas_streamlit(
                            uploaded_files=uploaded_files,
                            progress_widget=progress,
                            status_widget=status,
                            cambio_df=cambio_df
                        )
                        if df_final is not None and not df_final.empty:
                            st.success("Fluxo DUAS concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button(
                                    label="Baixar CSV (DUAS)",
                                    data=df_final.to_csv(index=False).encode("utf-8"),
                                    file_name="duas_consolidado.csv",
                                    mime="text/csv",
                                    use_container_width=True,
                                    key="duas_csv"
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="DUAS")
                                st.download_button(
                                    label="Baixar XLSX (DUAS)",
                                    data=xlsx_bytes,
                                    file_name="duas_consolidado.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="duas_xlsx"
                                )
                        else:
                            st.warning("Nenhuma tabela válida encontrada nos PDFs para o fluxo DUAS.")

                elif acao == "percepciones":
                    if not PERC_AVAILABLE:
                        st.error("Percepciones indisponível: confira dependências e `services/percepcion_service.py`.")
                    else:
                        df_final = process_percepcion_streamlit(
                            uploaded_files=uploaded_files,
                            progress_widget=progress,
                            status_widget=status,
                        )
                        if df_final is not None and not df_final.empty:
                            st.success("Percepciones concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button(
                                    label="Baixar CSV (Percepciones)",
                                    data=df_final.to_csv(index=False).encode("utf-8"),
                                    file_name="percepciones.csv",
                                    mime="text/csv",
                                    use_container_width=True,
                                    key="percepciones_csv"
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Percepciones")
                                st.download_button(
                                    label="Baixar XLSX (Percepciones)",
                                    data=xlsx_bytes,
                                    file_name="percepciones.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="percepciones_xlsx"
                                )
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Percepciones.")

                elif acao == "externos":
                    if not EXTERNOS_AVAILABLE:
                        st.error("Externos indisponível: confira dependências e `services/externos_service.py`.")
                    else:
                        cambio_df = st.session_state.get("tasa_df")
                        df_final = process_externos_streamlit(
                            uploaded_files=uploaded_files,
                            progress_widget=progress,
                            status_widget=status,
                            cambio_df=cambio_df,
                        )
                        if df_final is not None and not df_final.empty:
                            st.success("Externos concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button(
                                    label="Baixar CSV (Externos)",
                                    data=df_final.to_csv(index=False).encode("utf-8"),
                                    file_name="externos.csv",
                                    mime="text/csv",
                                    use_container_width=True,
                                    key="externos_csv"
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes_externos_duas_abas(
                                    df_normal=df_final,
                                    sheet_normal="Externos",
                                    sheet_spaced="Externos_Espacado",
                                    blank_rows=3,
                                )
                                st.download_button(
                                    label="Baixar XLSX (Externos)",
                                    data=xlsx_bytes,
                                    file_name="externos.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="externos_xlsx"
                                )
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Externos.")

                elif acao == "gastos":
                    if not ADICIONALES_AVAILABLE:
                        st.error("Gastos Adicionales indisponível: confira dependências e `services/adicionales_service.py`.")
                    else:
                        cambio_df = st.session_state.get("tasa_df")
                        df_final = process_adicionales_streamlit(
                            uploaded_files=uploaded_files,
                            progress_widget=progress,
                            status_widget=status,
                            cambio_df=cambio_df,
                        )
                        if df_final is not None and not df_final.empty:
                            st.success("Gastos Adicionales concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button(
                                    label="Baixar CSV (Adicionales)",
                                    data=df_final.to_csv(index=False).encode("utf-8"),
                                    file_name="gastos_adicionales.csv",
                                    mime="text/csv",
                                    use_container_width=True,
                                    key="adicionales_csv"
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Adicionales")
                                st.download_button(
                                    label="Baixar XLSX (Adicionales)",
                                    data=xlsx_bytes,
                                    file_name="gastos_adicionales.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="adicionales_xlsx"
                                )
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Gastos Adicionales.")

    # -------------------------
    # 🌐 Tasa SUNAT (renderiza SEMPRE, independente do tab1)
    # -------------------------
    with tab2:
        st.write("Baixar e consolidar Tasa (SUNAT) direto do site oficial.")
        anos = st.multiselect(
            "Anos",
            ["2024", "2025", "2026"],
            default=["2024", "2025", "2026"]
        )
        if st.button("Atualizar Tasa", key="tasa_update"):
            status = st.empty()
            pbar = st.progress(0, text="Iniciando...")
            df = atualizar_dataframe_tasa(
                anos=anos, progress_widget=pbar, status_widget=status
            )
            if df is not None and not df.empty:
                st.session_state.tasa_df = df.copy()
                st.success("Tasa consolidada com sucesso (armazenada para uso no DUAS/Externos).")
                st.dataframe(df.head(30), use_container_width=True)
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        label="Baixar CSV",
                        data=df.to_csv(index=False).encode("utf-8"),
                        file_name="tasa_consolidada.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key="tasa_csv"
                    )
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes(df, sheet_name="Tasa")
                    st.download_button(
                        label="Baixar XLSX",
                        data=xlsx_bytes,
                        file_name="tasa_consolidada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="tasa_xlsx"
                    )
            else:
                st.warning("Não foi possível obter dados da Tasa. Verifique credenciais/token/cookie.")

    # -------------------------
    # 📁 Arquivo Sharepoint (renderiza SEMPRE, independente do tab3)
    # -------------------------
    with tab3:
        st.subheader("📁 Arquivo Sharepoint")
        st.caption("Carregue um arquivo Excel para leitura da aba 'all'.")
        uploaded_excel = st.file_uploader(
            "Carregar Arquivo",
            type=["xlsx", "xls"],
            key="sharepoint_excel_uploader"
        )
        if uploaded_excel:
            try:
                df_all = pd.read_excel(
                    uploaded_excel,
                    sheet_name="all",
                    header=0,
                    usecols="A:Z",
                    nrows=20000,
                    engine="openpyxl"
                )
                from services.sharepoint_utils import ajustar_sharepoint_df
                df_all = ajustar_sharepoint_df(df_all)
                st.session_state["sharepoint_df"] = df_all
                st.success("✔️ DataFrame atualizado")
                st.dataframe(
                    df_all,
                    use_container_width=True,
                    height=500
                )

                st.subheader("⬇️ Downloads do Arquivo SharePoint")
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        label="Baixar CSV (SharePoint)",
                        data=df_all.to_csv(index=False).encode("utf-8"),
                        file_name="sharepoint_all.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key="sharepoint_csv"
                    )
                with col_xlsx:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df_all.to_excel(writer, index=False, sheet_name="SharePoint")
                        ws = writer.book["SharePoint"]
                        _autofit_worksheet(ws)
                    buffer.seek(0)
                    st.download_button(
                        label="Baixar XLSX (SharePoint)",
                        data=buffer,
                        file_name="sharepoint_all.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="sharepoint_xlsx"
                    )
            except ValueError:
                st.error("❌ A aba 'all' não foi encontrada no arquivo Excel.")
            except Exception as e:
                st.error("❌ Erro ao processar o arquivo Excel.")
                st.exception(e)

    with tab4:
        downloads_page.render()

    with tab5:
        st.subheader("📝 Transformar .prn")
        st.caption("Selecione um fluxo abaixo e carregue o Excel (primeira aba = Carga Financeira, segunda aba = Carga Contábil).")

        if "prn_flow" not in st.session_state:
            st.session_state.prn_flow = None  # "externos", "duas", "gastos"

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("Duas", key="prn_duas_select", use_container_width=True):
                st.session_state.prn_flow = "duas"
        with c2:
            if st.button("Externos", key="prn_externos_select", use_container_width=True):
                st.session_state.prn_flow = "externos"
        with c3:
            if st.button("Gastos Adicionales", key="prn_gastos_select", use_container_width=True):
                st.session_state.prn_flow = "gastos"

        st.divider()

        flow = st.session_state.prn_flow

        if flow == "externos":
            st.info("Fluxo **Externos** selecionado. Carregue o arquivo Excel.")
            uploaded_xl = st.file_uploader(
                "Carregar Excel (.xlsx ou .xls) para gerar PRN",
                type=["xlsx", "xls"],
                key="prn_externos_upl",
                accept_multiple_files=False,
                help="A 1ª aba será usada para 'Externos.prn' (Carga_Financeira) e a 2ª para 'aexternos.prn' (Carga_Contabil)."
            )

            if uploaded_xl is not None:
                colg1, colg2 = st.columns(2)

                with colg1:
                    if st.button("Gerar Externos.prn", key="gen_externos_prn1", type="primary", use_container_width=True):
                        try:
                            prn_bytes = gerar_externos_prn_primeira_aba(uploaded_xl)
                            st.success("Arquivo **Externos.prn** gerado!")
                            st.download_button(
                                "Baixar Externos.prn",
                                data=prn_bytes,
                                file_name="Externos.prn",
                                mime="text/plain",
                                use_container_width=True,
                                key="dl_externos_prn1"
                            )
                        except Exception as e:
                            st.error("Falha ao gerar Externos.prn")
                            st.exception(e)

                with colg2:
                    if st.button("Gerar aexternos.prn", key="gen_externos_prn2", type="secondary", use_container_width=True):
                        try:
                            prn_bytes2 = gerar_externos_prn_segunda_aba(uploaded_xl)
                            st.success("Arquivo **aexternos.prn** gerado!")
                            st.download_button(
                                "Baixar aexternos.prn",
                                data=prn_bytes2,
                                file_name="aexternos.prn",
                                mime="text/plain",
                                use_container_width=True,
                                key="dl_externos_prn2"
                            )
                        except Exception as e:
                            st.error("Falha ao gerar aexternos.prn")
                            st.exception(e)

        elif flow == "duas":
            st.info("Fluxo **Duas** selecionado. (Em breve: lógica específica para gerar PRN a partir do Excel.)")
            st.caption("Se quiser, já me passe a macro/algoritmo e eu programo aqui.")

        elif flow == "gastos":
            st.info("Fluxo **Gastos Adicionales** selecionado. Carregue o arquivo Excel.")
            uploaded_xl_g = st.file_uploader(
                "Carregar Excel (.xlsx ou .xls) para gerar PRN",
                type=["xlsx", "xls"],
                key="prn_gastos_upl",
                accept_multiple_files=False,
                help="1ª aba -> Adicionales.prn / ZIP (PRNs por linha) | 2ª aba -> AAdicionales.prn"
            )

            if uploaded_xl_g is not None:
                cga1, cga2, cga3 = st.columns(3)

                with cga1:
                    if st.button("Gerar Adicionales.prn", key="gen_adic_prn1", type="primary", use_container_width=True):
                        try:
                            prn_bytes_adic = gerar_adicionales_prn_primeira_aba(uploaded_xl_g)
                            st.success("Arquivo **Adicionales.prn** gerado!")
                            st.download_button(
                                "Baixar Adicionales.prn",
                                data=prn_bytes_adic,
                                file_name="Adicionales.prn",
                                mime="text/plain",
                                use_container_width=True,
                                key="dl_adic_prn1"
                            )
                        except Exception as e:
                            st.error("Falha ao gerar Adicionales.prn")
                            st.exception(e)

                with cga2:
                    if st.button("Gerar ZIP (PRNs por linha)", key="gen_adic_zip", type="secondary", use_container_width=True):
                        try:
                            zip_bytes, zip_name = gerar_adicionales_zip_primeira_aba(uploaded_xl_g, zip_name="Adicionales_PRNs.zip")
                            st.success("Arquivo **ZIP** com PRNs individuais gerado!")
                            st.download_button(
                                "Baixar Adicionales_PRNs.zip",
                                data=zip_bytes,
                                file_name=zip_name,
                                mime="application/zip",
                                use_container_width=True,
                                key="dl_adic_zip"
                            )
                        except Exception as e:
                            st.error("Falha ao gerar ZIP com PRNs individuais")
                            st.exception(e)

                with cga3:
                    if st.button("Gerar AAdicionales.prn", key="gen_adic_prn2", use_container_width=True):
                        try:
                            prn_bytes_aadic = gerar_adicionales_prn_segunda_aba(uploaded_xl_g)
                            st.success("Arquivo **AAdicionales.prn** gerado!")
                            st.download_button(
                                "Baixar AAdicionales.prn",
                                data=prn_bytes_aadic,
                                file_name="AAdicionales.prn",
                                mime="text/plain",
                                use_container_width=True,
                                key="dl_adic_prn2"
                            )
                        except Exception as e:
                            st.error("Falha ao gerar AAdicionales.prn")
                            st.exception(e)

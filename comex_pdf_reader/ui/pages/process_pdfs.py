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
        st.exception(DUAS_ERR)

# --- NOVO: Import protegido do Externos ---
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
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# ============================
# CONFIG: Espelhar larguras PRN nos XLSX
# ============================
USE_PRN_WIDTHS = True  # <- altere para False se quiser voltar ao autoajuste

# Larguras fixas (em "caracteres", aproximadas ao Excel) - iguais aos PRN
PRN_WIDTHS_1 = [10, 25, 6, 6, 6, 16, 16, 2, 5, 16, 3, 2, 30, 6, 3, 3, 8, 3, 6, 4, 16, 16, 3, 6]  # 24 colunas
PRN_WIDTHS_2 = [6, 3, 3, 8, 3, 16, 16, 2, 30, 6, 15, 20, 5]                                      # 13 colunas

def set_fixed_widths(ws, widths, start_col: int = 1, add_excel_padding: bool = True):
    """
    Define larguras fixas por coluna.
    Quando add_excel_padding=True, soma ~0.71 para que o Excel exiba o mesmo número do PRN
    na caixa 'Column width' (ex.: gravar 10.71 para a UI mostrar 10.00).
    """
    PADDING = 0.71 if add_excel_padding else 0.0
    for i, w in enumerate(widths, start=start_col):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = float(w) + PADDING

def _autofit_worksheet(ws, font_padding: float = 1.2, min_width: float = 8.0, max_width: float = 60.0):
    if ws.max_column is None or ws.max_row is None:
        return
    for col_idx, col in enumerate(
        ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column),
        start=1
    ):
        header = col[0].value if col and col[0] is not None else ""
        max_len = len(str(header)) if header is not None else 0
        for cell in col[1:]:
            val = cell.value
            if val is None:
                continue
            # para cálculo de comprimento apenas
            text = f"{val:.6g}" if isinstance(val, float) else str(val)
            max_len = max(max_len, len(text))
        width = min(max(max_len * font_padding, min_width), max_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def header_paint(ws):
    BLUE = "FF0077B6"
    WHITE = "FFFFFFFF"
    fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
    font_white_bold = Font(color=WHITE, bold=True)

    exact_headers = {
        "source_file","Proveedor","Proveedor Iscala","Factura","Tipo Doc","Cód. de Autorización",
        "Tipo de Factura","Fecha de Emisión","Moneda","Cod. Moneda","Amount","Tasa","Cuenta",
        "Error","R.U.C","Op. Gravada","COD PROVEEDOR","Declaracion","Fecha","Ad_Valorem",
        "Imp_Prom_Municipal","Imp_Gene_a_las_Ventas","IGV","Percepcion","PEC","COD Moneda",
        "Source_File","No_Liquidacion","CDA","Monto","COD MONEDA","Lineaabajo",
    }
    for cell in ws[1]:
        if cell.value is None:
            continue
        header_text = str(cell.value).strip()
        if header_text in exact_headers:
            cell.fill = fill_blue
            cell.font = font_white_bold

def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Tasa") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        # Se quisermos espelhar PRN quando a contagem for 24 ou 13, aplicamos; senão, autoajuste
        if USE_PRN_WIDTHS:
            if df.shape[1] == 24:
                set_fixed_widths(ws, PRN_WIDTHS_1, start_col=1)
            elif df.shape[1] == 13:
                set_fixed_widths(ws, PRN_WIDTHS_2, start_col=1)
            else:
                _autofit_worksheet(ws)
        else:
            _autofit_worksheet(ws)
        header_paint(ws)
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# DF com linhas em branco entre cada registro (para XLSX de Externos)
# -----------------------------------------------------------------------------
def df_with_blank_spacers(df: pd.DataFrame, blank_rows: int = 3) -> pd.DataFrame:
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

def to_xlsx_bytes_externos_duas_abas(
    df_normal: pd.DataFrame,
    sheet_normal: str = "Externos",
    sheet_spaced: str = "Externos_Espacado",
    blank_rows: int = 3
) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Aba 1 (normal)
        df_normal.to_excel(writer, index=False, sheet_name=sheet_normal)
        ws1 = writer.book[sheet_normal]
        if USE_PRN_WIDTHS and df_normal.shape[1] == 24:
            set_fixed_widths(ws1, PRN_WIDTHS_1, start_col=1)
        else:
            _autofit_worksheet(ws1)
        header_paint(ws1)

        # Aba 2 (espacada)
        df_spaced = df_with_blank_spacers(df_normal, blank_rows=blank_rows)
        df_spaced.to_excel(writer, index=False, sheet_name=sheet_spaced)
        ws2 = writer.book[sheet_spaced]
        if USE_PRN_WIDTHS and df_spaced.shape[1] == 24:
            set_fixed_widths(ws2, PRN_WIDTHS_1, start_col=1)
        else:
            _autofit_worksheet(ws2)
        header_paint(ws2)
    buffer.seek(0)
    return buffer.getvalue()

ACTIONS = {"externos": "Externos","gastos": "Gastos Adicionales","duas": "Duas","percepciones": "Percepciones"}

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

# ================== HELPERS (PRN) ==================
import math
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

def _to_str(x):
    if x is None:
        return ""
    if isinstance(x, float) and math.isnan(x):
        return ""
    s = str(x)
    return "" if s.strip() in {"nan", "NaN"} else s

def _format_decimal_2_dot(value):
    if value is None:
        return ""
    txt = str(value).strip()
    if txt == "":
        return ""
    txt_norm = txt.replace(",", ".")
    try:
        d = Decimal(txt_norm)
    except (InvalidOperation, ValueError):
        return txt
    d2 = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{d2}"

def _fixed_width_line(values, widths, fmt=None):
    out = []
    for idx, (v, w) in enumerate(zip(values, widths)):
        sv = fmt(idx, v) if fmt else _to_str(v)
        sv = sv[:w]
        out.append(sv + (" " * max(0, w - len(sv))))
    return "".join(out)

def _df_to_prn_bytes(rows_values, widths, encoding="cp1252", fmt=None):
    lines = [_fixed_width_line(vals, widths, fmt=fmt) for vals in rows_values]
    text = "\r\n".join(lines) + "\r\n"
    return text.encode(encoding, errors="replace")

# ------------------- EXTERNOS: 1ª ABA -> PRN -------------------
def gerar_externos_prn_primeira_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        row_idx = r - 2
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df.index) and 0 <= col_idx < len(df.columns):
                return df.iloc[row_idx, col_idx]
        except Exception:
            return ""
        return ""

    widths = PRN_WIDTHS_1[:]  # 24 colunas

    rows_values = []
    for r in range(3, 1501, 4):
        val_c = _to_str(get_cell(r, 3))
        if val_c != "":
            row_vals = [get_cell(r, c) for c in range(3, 27)]
            rows_values.append(row_vals)

    DEC2_COLS = {5, 6, 9, 20, 21}  # F, G, J, U, V (0-based)
    def fmt(col_idx, value):
        if col_idx in DEC2_COLS:
            return _format_decimal_2_dot(value)
        return _to_str(value)

    prn_bytes = _df_to_prn_bytes(rows_values, widths, encoding="cp1252", fmt=fmt)
    return prn_bytes

# ------------------- EXTERNOS: 2ª ABA -> PRN -------------------
def gerar_externos_prn_segunda_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df2 = pd.read_excel(xls_file, sheet_name=1, header=0, dtype=str, engine=engine)

    def get_cell2(r, c):
        row_idx = r - 2
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df2.index) and 0 <= col_idx < len(df2.columns):
                return df2.iloc[row_idx, col_idx]
        except Exception:
            return ""
        return ""

    linha_limite = 0
    for r in range(2, 46001):
        val_b = get_cell2(r, 2)
        if (val_b is None) or (str(val_b).strip() in {"#N/A", "#N/D"}) or (pd.isna(val_b)):
            linha_limite = r - 4
            break
    if linha_limite <= 0:
        linha_limite = 1496

    rows_raw = []
    for r in range(2, max(2, linha_limite) + 1):
        row_vals = [get_cell2(r, c) for c in range(2, 15)]  # 13 colunas
        rows_raw.append(row_vals)

    rows_clean = []
    for vals in rows_raw:
        d_val = _to_str(vals[3])
        if d_val.strip() in {"0", "0.0"}:
            vals[3] = ""
        f_val = _to_str(vals[5]).strip()
        if f_val in {"", "0", "0.0"}:
            continue
        rows_clean.append(vals)

    widths2 = PRN_WIDTHS_2[:]  # 13 colunas

    DEC2_COLS = {5}  # apenas F (0-based)
    def fmt(col_idx, value):
        if col_idx in DEC2_COLS:
            return _format_decimal_2_dot(value)
        return _to_str(value)

    prn_bytes = _df_to_prn_bytes(rows_clean, widths2, encoding="cp1252", fmt=fmt)
    return prn_bytes

# ------------------- ADICIONALES: 1ª ABA -> PRN -------------------
def gerar_adicionales_prn_primeira_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        row_idx = r - 2
        col_idx = c - 1
        try:
            if 0 <= row_idx < len(df.index) and 0 <= col_idx < len(df.columns):
                return df.iloc[row_idx, col_idx]
        except Exception:
            return ""
        return ""

    widths = PRN_WIDTHS_1[:]  # 24 colunas

    rows_values = []
    for r in range(3, 1501, 4):
        val_c = _to_str(get_cell(r, 3))
        if val_c != "":
            row_vals = [get_cell(r, c) for c in range(3, 27)]
            rows_values.append(row_vals)

    DEC2_COLS = {5, 6, 9, 20, 21}
    def fmt(col_idx, value):
        if col_idx in DEC2_COLS:
            return _format_decimal_2_dot(value)
        return _to_str(value)

    prn_bytes = _df_to_prn_bytes(rows_values, widths, encoding="cp1252", fmt=fmt)
    return prn_bytes

# ------------------- ADICIONALES: 1ª ABA -> ZIP (PRN por linha) -------------------
def gerar_adicionales_zip_primeira_aba(xls_file, zip_name="Adicionales_PRNs.zip"):
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
                return df.iloc[row_idx, col_idx]
        except Exception:
            return ""
        return ""

    widths = PRN_WIDTHS_1[:]  # 24 colunas

    buffer_zip = BytesIO()
    with ZipFile(buffer_zip, mode="w", compression=ZIP_DEFLATED) as zf:
        seq = 1
        DEC2_COLS = {5, 6, 9, 20, 21}
        def fmt(col_idx, value):
            if col_idx in DEC2_COLS:
                return _format_decimal_2_dot(value)
            return _to_str(value)

        for r in range(3, 1501, 4):
            val_c = _to_str(get_cell(r, 3))
            if val_c == "":
                continue
            row_vals = [get_cell(r, c) for c in range(3, 27)]
            prn_bytes = _df_to_prn_bytes([row_vals], widths, encoding="cp1252", fmt=fmt)

            safe_prefix = (val_c or "linha").replace("\\", "_").replace("/", "_").replace(" ", "")
            filename = f"{safe_prefix}_{seq}.prn"
            zf.writestr(filename, prn_bytes)
            seq += 1

    buffer_zip.seek(0)
    return buffer_zip.getvalue(), zip_name

# ------------------- ADICIONALES: 2ª ABA -> PRN -------------------
def gerar_adicionales_prn_segunda_aba(xls_file):
    return gerar_externos_prn_segunda_aba(xls_file)  # mesma regra

# ================== HELPERS XLSX (novos) ==================

def _rows_to_xlsx_bytes(rows, headers, sheet_name, decimal_cols_idx=None):
    """Converte 'rows' (lista de listas) em XLSX com headers, aplicando 2 casas nas colunas indicadas."""
    decimal_cols_idx = decimal_cols_idx or set()
    # normaliza e aplica 2 casas onde necessário
    rows_fmt = []
    for r in rows:
        r2 = []
        for i, v in enumerate(r):
            if i in decimal_cols_idx:
                r2.append(_format_decimal_2_dot(v))
            else:
                r2.append(_to_str(v))
        rows_fmt.append(r2)

    df_out = pd.DataFrame(rows_fmt, columns=headers)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        # Aplica PRN widths quando possível
        if USE_PRN_WIDTHS:
            if len(headers) == 24:
                set_fixed_widths(ws, PRN_WIDTHS_1, start_col=1)
            elif len(headers) == 13:
                set_fixed_widths(ws, PRN_WIDTHS_2, start_col=1)
            else:
                _autofit_worksheet(ws)
        else:
            _autofit_worksheet(ws)
        header_paint(ws)
    buffer.seek(0)
    return buffer.getvalue()

# EXTERNOS 1ª aba -> XLSX
def gerar_externos_xlsx_primeira_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        ri = r - 2
        ci = c - 1
        try:
            if 0 <= ri < len(df.index) and 0 <= ci < len(df.columns):
                return df.iloc[ri, ci]
        except Exception:
            return ""
        return ""

    rows = []
    for r in range(3, 1501, 4):
        if _to_str(get_cell(r, 3)) != "":
            rows.append([get_cell(r, c) for c in range(3, 27)])  # 24 colunas

    headers = [f"Col_{chr(65+i)}" for i in range(24)]  # A..X
    decimal_cols = {5, 6, 9, 20, 21}  # F,G,J,U,V
    return _rows_to_xlsx_bytes(rows, headers, "Externos", decimal_cols)

# EXTERNOS 2ª aba -> XLSX
def gerar_externos_xlsx_segunda_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df2 = pd.read_excel(xls_file, sheet_name=1, header=0, dtype=str, engine=engine)

    def get_cell2(r, c):
        ri = r - 2
        ci = c - 1
        try:
            if 0 <= ri < len(df2.index) and 0 <= ci < len(df2.columns):
                return df2.iloc[ri, ci]
        except Exception:
            return ""
        return ""

    linha_limite = 0
    for r in range(2, 46001):
        val_b = get_cell2(r, 2)
        if (val_b is None) or (str(val_b).strip() in {"#N/A", "#N/D"}) or (pd.isna(val_b)):
            linha_limite = r - 4
            break
    if linha_limite <= 0:
        linha_limite = 1496

    rows_raw = []
    for r in range(2, max(2, linha_limite) + 1):
        rows_raw.append([get_cell2(r, c) for c in range(2, 15)])  # 13 colunas

    rows_clean = []
    for vals in rows_raw:
        d_val = _to_str(vals[3])
        if d_val.strip() in {"0", "0.0"}:
            vals[3] = ""
        f_val = _to_str(vals[5]).strip()
        if f_val in {"", "0", "0.0"}:
            continue
        rows_clean.append(vals)

    headers = [f"Col_{chr(65+i)}" for i in range(13)]  # A..M
    decimal_cols = {5}  # F
    return _rows_to_xlsx_bytes(rows_clean, headers, "aexternos", decimal_cols)

# ADICIONALES 1ª aba -> XLSX
def gerar_adicionales_xlsx_primeira_aba(xls_file):
    return gerar_externos_xlsx_primeira_aba(xls_file)

# ADICIONALES 2ª aba -> XLSX
def gerar_adicionales_xlsx_segunda_aba(xls_file):
    return gerar_externos_xlsx_segunda_aba(xls_file)

# ======================= DUAS - 1ª ABA → PRN =======================
def gerar_duas_prn_primeira_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        ri = r - 2
        ci = c - 1
        try:
            if 0 <= ri < len(df.index) and 0 <= ci < len(df.columns):
                return df.iloc[ri, ci]
        except Exception:
            return ""
        return ""

    widths = PRN_WIDTHS_1[:]  # 24 colunas

    rows = []
    for r in range(3, 1501, 4):
        if _to_str(get_cell(r, 3)) != "":
            rows.append([get_cell(r, c) for c in range(3, 27)])

    DEC2_COLS = {5, 6, 9, 20, 21}
    def fmt(col_idx, value):
        return _format_decimal_2_dot(value) if col_idx in DEC2_COLS else _to_str(value)

    return _df_to_prn_bytes(rows, widths, encoding="cp1252", fmt=fmt)

# ======================= DUAS - 2ª ABA → PRN =======================
def gerar_duas_prn_segunda_aba(xls_file):
    # Mesma regra usada para Externos/Adicionales 2ª aba
    return gerar_externos_prn_segunda_aba(xls_file)

# ======================= DUAS – ZIP (PRNs por linha – 1ª aba) =======================
def gerar_duas_zip_primeira_aba(xls_file, zip_name="Duas_PRNs.zip"):
    # Reutiliza a mesma lógica do ZIP de Adicionales, mudando apenas o nome do arquivo
    return gerar_adicionales_zip_primeira_aba(xls_file, zip_name=zip_name)

# ======================= DUAS - 1ª ABA → XLSX =======================
def gerar_duas_xlsx_primeira_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(xls_file, sheet_name=0, header=0, dtype=str, engine=engine)

    def get_cell(r, c):
        ri = r - 2
        ci = c - 1
        try:
            if 0 <= ri < len(df.index) and 0 <= ci < len(df.columns):
                return df.iloc[ri, ci]
        except Exception:
            return ""
        return ""

    rows = []
    for r in range(3, 1501, 4):
        if _to_str(get_cell(r, 3)) != "":
            rows.append([get_cell(r, c) for c in range(3, 27)])  # 24 colunas

    headers = [f"Col_{chr(65+i)}" for i in range(24)]  # A..X
    decimal_cols = {5, 6, 9, 20, 21}  # F,G,J,U,V
    return _rows_to_xlsx_bytes(rows, headers, "Duas", decimal_cols)

# ======================= DUAS - 2ª ABA → XLSX =======================
def gerar_duas_xlsx_segunda_aba(xls_file):
    name = getattr(xls_file, "name", "").lower()
    engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
    df2 = pd.read_excel(xls_file, sheet_name=1, header=0, dtype=str, engine=engine)

    def get_cell2(r, c):
        ri = r - 2
        ci = c - 1
        try:
            if 0 <= ri < len(df2.index) and 0 <= ci < len(df2.columns):
                return df2.iloc[ri, ci]
        except Exception:
            return ""
        return ""

    linha_limite = 0
    for r in range(2, 46001):
        val_b = get_cell2(r, 2)
        if (val_b is None) or (str(val_b).strip() in {"#N/A", "#N/D"}) or (pd.isna(val_b)):
            linha_limite = r - 4
            break
    if linha_limite <= 0:
        linha_limite = 1496

    rows_raw = []
    for r in range(2, max(2, linha_limite) + 1):
        rows_raw.append([get_cell2(r, c) for c in range(2, 15)])  # 13 colunas

    rows_clean = []
    for vals in rows_raw:
        d_val = _to_str(vals[3])
        if d_val.strip() in {"0", "0.0"}:
            vals[3] = ""
        f_val = _to_str(vals[5]).strip()
        if f_val in {"", "0", "0.0"}:
            continue
        rows_clean.append(vals)

    headers = [f"Col_{chr(65+i)}" for i in range(13)]  # A..M
    decimal_cols = {5}  # F
    return _rows_to_xlsx_bytes(rows_clean, headers, "ADuas", decimal_cols)

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
            st.warning("O módulo **DUAS** não pôde ser carregado. Verifique `services/duas_service.py` e dependências (ex.: `pdfplumber`).")

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
                run_clicked = st.button("▶️ Executar", key="action_run", type="primary", use_container_width=True, disabled=not uploaded_files)
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
                        df_final = process_duas_streamlit(uploaded_files=uploaded_files, progress_widget=progress, status_widget=status, cambio_df=cambio_df)
                        if df_final is not None and not df_final.empty:
                            st.success("Fluxo DUAS concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button("Baixar CSV (DUAS)", data=df_final.to_csv(index=False).encode("utf-8"), file_name="duas_consolidado.csv", mime="text/csv", use_container_width=True, key="duas_csv")
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="DUAS")
                                st.download_button("Baixar XLSX (DUAS)", data=xlsx_bytes, file_name="duas_consolidado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="duas_xlsx")
                        else:
                            st.warning("Nenhuma tabela válida encontrada nos PDFs para o fluxo DUAS.")

                elif acao == "percepciones":
                    if not PERC_AVAILABLE:
                        st.error("Percepciones indisponível: confira dependências e `services/percepcion_service.py`.")
                    else:
                        df_final = process_percepcion_streamlit(uploaded_files=uploaded_files, progress_widget=progress, status_widget=status)
                        if df_final is not None and not df_final.empty:
                            st.success("Percepciones concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button("Baixar CSV (Percepciones)", data=df_final.to_csv(index=False).encode("utf-8"), file_name="percepciones.csv", mime="text/csv", use_container_width=True, key="percepciones_csv")
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Percepciones")
                                st.download_button("Baixar XLSX (Percepciones)", data=xlsx_bytes, file_name="percepciones.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="percepciones_xlsx")
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Percepciones.")

                elif acao == "externos":
                    if not EXTERNOS_AVAILABLE:
                        st.error("Externos indisponível: confira dependências e `services/externos_service.py`.")
                    else:
                        cambio_df = st.session_state.get("tasa_df")
                        df_final = process_externos_streamlit(uploaded_files=uploaded_files, progress_widget=progress, status_widget=status, cambio_df=cambio_df)
                        if df_final is not None and not df_final.empty:
                            st.success("Externos concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button("Baixar CSV (Externos)", data=df_final.to_csv(index=False).encode("utf-8"), file_name="externos.csv", mime="text/csv", use_container_width=True, key="externos_csv")
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes_externos_duas_abas(df_normal=df_final, sheet_normal="Externos", sheet_spaced="Externos_Espacado", blank_rows=3)
                                st.download_button("Baixar XLSX (Externos)", data=xlsx_bytes, file_name="externos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="externos_xlsx")
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Externos.")

                elif acao == "gastos":
                    if not ADICIONALES_AVAILABLE:
                        st.error("Gastos Adicionales indisponível: confira dependências e `services/adicionales_service.py`.")
                    else:
                        cambio_df = st.session_state.get("tasa_df")
                        df_final = process_adicionales_streamlit(uploaded_files=uploaded_files, progress_widget=progress, status_widget=status, cambio_df=cambio_df)
                        if df_final is not None and not df_final.empty:
                            st.success("Gastos Adicionales concluído!")
                            st.dataframe(df_final.head(50), use_container_width=True)
                            col_csv, col_xlsx = st.columns(2)
                            with col_csv:
                                st.download_button("Baixar CSV (Adicionales)", data=df_final.to_csv(index=False).encode("utf-8"), file_name="gastos_adicionales.csv", mime="text/csv", use_container_width=True, key="adicionales_csv")
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Adicionales")
                                st.download_button("Baixar XLSX (Adicionales)", data=xlsx_bytes, file_name="gastos_adicionales.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="adicionales_xlsx")
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Gastos Adicionales.")

    # -------------------------
    # 🌐 Tasa SUNAT
    # -------------------------
    with tab2:
        st.write("Baixar e consolidar Tasa (SUNAT) direto do site oficial.")
        anos = st.multiselect("Anos", ["2024", "2025", "2026"], default=["2024", "2025", "2026"])
        if st.button("Atualizar Tasa", key="tasa_update"):
            status = st.empty()
            pbar = st.progress(0, text="Iniciando...")
            df = atualizar_dataframe_tasa(anos=anos, progress_widget=pbar, status_widget=status)
            if df is not None and not df.empty:
                st.session_state.tasa_df = df.copy()
                st.success("Tasa consolidada com sucesso (armazenada para uso no DUAS/Externos).")
                st.dataframe(df.head(30), use_container_width=True)
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button("Baixar CSV", data=df.to_csv(index=False).encode("utf-8"), file_name="tasa_consolidada.csv", mime="text/csv", use_container_width=True, key="tasa_csv")
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes(df, sheet_name="Tasa")
                    st.download_button("Baixar XLSX", data=xlsx_bytes, file_name="tasa_consolidada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="tasa_xlsx")
            else:
                st.warning("Não foi possível obter dados da Tasa. Verifique credenciais/token/cookie.")

    # -------------------------
    # 📁 Arquivo Sharepoint
    # -------------------------
    with tab3:
        st.subheader("📁 Arquivo Sharepoint")
        st.caption("Carregue um arquivo Excel para leitura da aba 'all'.")
        uploaded_excel = st.file_uploader("Carregar Arquivo", type=["xlsx", "xls"], key="sharepoint_excel_uploader")
        if uploaded_excel:
            try:
                df_all = pd.read_excel(uploaded_excel, sheet_name="all", header=0, usecols="A:Z", nrows=20000, engine="openpyxl")
                from services.sharepoint_utils import ajustar_sharepoint_df
                df_all = ajustar_sharepoint_df(df_all)
                st.session_state["sharepoint_df"] = df_all
                st.success("✔️ DataFrame atualizado")
                st.dataframe(df_all, use_container_width=True, height=500)

                st.subheader("⬇️ Downloads do Arquivo SharePoint")
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button("Baixar CSV (SharePoint)", data=df_all.to_csv(index=False).encode("utf-8"), file_name="sharepoint_all.csv", mime="text/csv", use_container_width=True, key="sharepoint_csv")
                with col_xlsx:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df_all.to_excel(writer, index=False, sheet_name="SharePoint")
                        ws = writer.book["SharePoint"]
                        if USE_PRN_WIDTHS and df_all.shape[1] in (24, 13):
                            # só aplica PRN widths se bater com 24 ou 13; senão, autoajuste
                            set_fixed_widths(ws, PRN_WIDTHS_1 if df_all.shape[1] == 24 else PRN_WIDTHS_2, start_col=1)
                        else:
                            _autofit_worksheet(ws)
                    buffer.seek(0)
                    st.download_button("Baixar XLSX (SharePoint)", data=buffer, file_name="sharepoint_all.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="sharepoint_xlsx")
            except ValueError:
                st.error("❌ A aba 'all' não foi encontrada no arquivo Excel.")
            except Exception as e:
                st.error("❌ Erro ao processar o arquivo Excel.")
                st.exception(e)

    with tab4:
        downloads_page.render()

    # -------------------------
    # 📝 Transformar .prn
    # -------------------------
    with tab5:
        st.subheader("📝 Transformar .prn")
        st.caption("Selecione um fluxo abaixo e carregue o Excel (primeira aba = Carga Financeira, segunda aba = Carga Contábil).")

        if "prn_flow" not in st.session_state:
            st.session_state.prn_flow = None

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

        # -------------------- FLUXO EXTERNOS --------------------
        if flow == "externos":
            st.info("Fluxo **Externos** selecionado. Carregue o arquivo Excel.")
            uploaded_xl = st.file_uploader(
                "Carregar Excel (.xlsx ou .xls) para gerar PRN/XLSX",
                type=["xlsx", "xls"],
                key="prn_externos_upl",
                accept_multiple_files=False,
                help="1ª aba -> Externos.prn/.xlsx | 2ª aba -> aexternos.prn/.xlsx"
            )

            if uploaded_xl is not None:
                colg1, colg2 = st.columns(2)
                with colg1:
                    # PRN (1ª aba)
                    if st.button("Gerar Externos.prn", key="gen_externos_prn1", type="primary", use_container_width=True):
                        try:
                            prn_bytes = gerar_externos_prn_primeira_aba(uploaded_xl)
                            st.success("Arquivo **Externos.prn** gerado!")
                            st.download_button("Baixar Externos.prn", data=prn_bytes, file_name="Externos.prn", mime="text/plain", use_container_width=True, key="dl_externos_prn1")
                        except Exception as e:
                            st.error("Falha ao gerar Externos.prn")
                            st.exception(e)

                    # XLSX (1ª aba)
                    if st.button("Gerar Externos.xlsx", key="gen_externos_xlsx1", use_container_width=True):
                        try:
                            xlsx_bytes = gerar_externos_xlsx_primeira_aba(uploaded_xl)
                            st.success("Arquivo **Externos.xlsx** gerado!")
                            st.download_button("Baixar Externos.xlsx", data=xlsx_bytes, file_name="Externos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_externos_xlsx1")
                        except Exception as e:
                            st.error("Falha ao gerar Externos.xlsx")
                            st.exception(e)

                with colg2:
                    # PRN (2ª aba)
                    if st.button("Gerar aexternos.prn", key="gen_externos_prn2", type="secondary", use_container_width=True):
                        try:
                            prn_bytes2 = gerar_externos_prn_segunda_aba(uploaded_xl)
                            st.success("Arquivo **aexternos.prn** gerado!")
                            st.download_button("Baixar aexternos.prn", data=prn_bytes2, file_name="aexternos.prn", mime="text/plain", use_container_width=True, key="dl_externos_prn2")
                        except Exception as e:
                            st.error("Falha ao gerar aexternos.prn")
                            st.exception(e)

                    # XLSX (2ª aba)
                    if st.button("Gerar aexternos.xlsx", key="gen_externos_xlsx2", use_container_width=True):
                        try:
                            xlsx_bytes2 = gerar_externos_xlsx_segunda_aba(uploaded_xl)
                            st.success("Arquivo **aexternos.xlsx** gerado!")
                            st.download_button("Baixar aexternos.xlsx", data=xlsx_bytes2, file_name="aexternos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_externos_xlsx2")
                        except Exception as e:
                            st.error("Falha ao gerar aexternos.xlsx")
                            st.exception(e)

        # -------------------- FLUXO DUAS --------------------
        elif flow == "duas":
            st.info("Fluxo **DUAS** selecionado. Carregue o arquivo Excel.")
            uploaded_xl_d = st.file_uploader(
                "Carregar Excel (.xlsx ou .xls) para gerar PRN/XLSX",
                type=["xlsx", "xls"],
                key="prn_duas_upl",
                accept_multiple_files=False,
                help="1ª aba -> Duas.prn/.xlsx | 2ª aba -> ADuas.prn/.xlsx"
            )

            if uploaded_xl_d is not None:
                c1, c2 = st.columns(2)

                # -------- PRIMEIRA ABA --------
                with c1:
                    if st.button("Gerar Duas.prn", key="gen_duas_prn1", type="primary", use_container_width=True):
                        try:
                            prn_bytes = gerar_duas_prn_primeira_aba(uploaded_xl_d)
                            st.success("Arquivo **Duas.prn** gerado!")
                            st.download_button("Baixar Duas.prn", data=prn_bytes, file_name="Duas.prn",
                                               mime="text/plain", use_container_width=True, key="dl_duas_prn1")
                        except Exception as e:
                            st.error("Falha ao gerar Duas.prn")
                            st.exception(e)

                    # NOVO: ZIP PRNs por linha (1ª aba)
                    if st.button("Gerar ZIP (PRNs por linha)", key="gen_duas_zip", type="secondary", use_container_width=True):
                        try:
                            zip_bytes, zip_name = gerar_duas_zip_primeira_aba(uploaded_xl_d, zip_name="Duas_PRNs.zip")
                            st.success("Arquivo **ZIP** com PRNs individuais gerado!")
                            st.download_button("Baixar Duas_PRNs.zip", data=zip_bytes, file_name=zip_name,
                                               mime="application/zip", use_container_width=True, key="dl_duas_zip")
                        except Exception as e:
                            st.error("Falha ao gerar ZIP com PRNs individuais (DUAS)")
                            st.exception(e)

                    if st.button("Gerar Duas.xlsx", key="gen_duas_xlsx1", use_container_width=True):
                        try:
                            xlsx1 = gerar_duas_xlsx_primeira_aba(uploaded_xl_d)
                            st.success("Arquivo **Duas.xlsx** gerado!")
                            st.download_button("Baixar Duas.xlsx", data=xlsx1, file_name="Duas.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                               use_container_width=True, key="dl_duas_xlsx1")
                        except Exception as e:
                            st.error("Falha ao gerar Duas.xlsx")
                            st.exception(e)

                # -------- SEGUNDA ABA --------
                with c2:
                    if st.button("Gerar ADuas.prn", key="gen_duas_prn2", use_container_width=True):
                        try:
                            prn_bytes2 = gerar_duas_prn_segunda_aba(uploaded_xl_d)
                            st.success("Arquivo **ADuas.prn** gerado!")
                            st.download_button("Baixar ADuas.prn", data=prn_bytes2, file_name="ADuas.prn",
                                               mime="text/plain", use_container_width=True, key="dl_duas_prn2")
                        except Exception as e:
                            st.error("Falha ao gerar ADuas.prn")
                            st.exception(e)

                    if st.button("Gerar ADuas.xlsx", key="gen_duas_xlsx2", use_container_width=True):
                        try:
                            xlsx2 = gerar_duas_xlsx_segunda_aba(uploaded_xl_d)
                            st.success("Arquivo **ADuas.xlsx** gerado!")
                            st.download_button("Baixar ADuas.xlsx", data=xlsx2, file_name="ADuas.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                               use_container_width=True, key="dl_duas_xlsx2")
                        except Exception as e:
                            st.error("Falha ao gerar ADuas.xlsx")
                            st.exception(e)

        # -------------------- FLUXO GASTOS ADICIONALES --------------------
        elif flow == "gastos":
            st.info("Fluxo **Gastos Adicionales** selecionado. Carregue o arquivo Excel.")
            uploaded_xl_g = st.file_uploader(
                "Carregar Excel (.xlsx ou .xls) para gerar PRN/XLSX",
                type=["xlsx", "xls"],
                key="prn_gastos_upl",
                accept_multiple_files=False,
                help="1ª aba -> Adicionales.prn/.xlsx | 2ª aba -> AAdicionales.prn/.xlsx"
            )

            if uploaded_xl_g is not None:
                cga1, cga2, cga3 = st.columns(3)

                with cga1:
                    if st.button("Gerar Adicionales.prn", key="gen_adic_prn1", type="primary", use_container_width=True):
                        try:
                            prn_bytes_adic = gerar_adicionales_prn_primeira_aba(uploaded_xl_g)
                            st.success("Arquivo **Adicionales.prn** gerado!")
                            st.download_button("Baixar Adicionales.prn", data=prn_bytes_adic, file_name="Adicionales.prn", mime="text/plain", use_container_width=True, key="dl_adic_prn1")
                        except Exception as e:
                            st.error("Falha ao gerar Adicionales.prn")
                            st.exception(e)

                    # XLSX (1ª aba)
                    if st.button("Gerar Adicionales.xlsx", key="gen_adic_xlsx1", use_container_width=True):
                        try:
                            xlsx_adic_1 = gerar_adicionales_xlsx_primeira_aba(uploaded_xl_g)
                            st.success("Arquivo **Adicionales.xlsx** gerado!")
                            st.download_button("Baixar Adicionales.xlsx", data=xlsx_adic_1, file_name="Adicionales.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_adic_xlsx1")
                        except Exception as e:
                            st.error("Falha ao gerar Adicionales.xlsx")
                            st.exception(e)

                with cga2:
                    if st.button("Gerar ZIP (PRNs por linha)", key="gen_adic_zip", type="secondary", use_container_width=True):
                        try:
                            zip_bytes, zip_name = gerar_adicionales_zip_primeira_aba(uploaded_xl_g, zip_name="Adicionales_PRNs.zip")
                            st.success("Arquivo **ZIP** com PRNs individuais gerado!")
                            st.download_button("Baixar Adicionales_PRNs.zip", data=zip_bytes, file_name=zip_name, mime="application/zip", use_container_width=True, key="dl_adic_zip")
                        except Exception as e:
                            st.error("Falha ao gerar ZIP com PRNs individuais")
                            st.exception(e)

                with cga3:
                    if st.button("Gerar AAdicionales.prn", key="gen_adic_prn2", use_container_width=True):
                        try:
                            prn_bytes_aadic = gerar_adicionales_prn_segunda_aba(uploaded_xl_g)
                            st.success("Arquivo **AAdicionales.prn** gerado!")
                            st.download_button("Baixar AAdicionales.prn", data=prn_bytes_aadic, file_name="AAdicionales.prn", mime="text/plain", use_container_width=True, key="dl_adic_prn2")
                        except Exception as e:
                            st.error("Falha ao gerar AAdicionales.prn")
                            st.exception(e)

                    # XLSX (2ª aba)
                    if st.button("Gerar AAdicionales.xlsx", key="gen_adic_xlsx2", use_container_width=True):
                        try:
                            xlsx_adic_2 = gerar_adicionales_xlsx_segunda_aba(uploaded_xl_g)
                            st.success("Arquivo **AAdicionales.xlsx** gerado!")
                            st.download_button("Baixar AAdicionales.xlsx", data=xlsx_adic_2, file_name="AAdicionales.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_adic_xlsx2")
                        except Exception as e:
                            st.error("Falha ao gerar AAdicionales.xlsx")
                            st.exception(e)

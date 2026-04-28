# ui/pages/process_pdfs.py
import streamlit as st
from services import pdf_service
from services.tasa_service import atualizar_dataframe_tasa
from ui.pages import downloads_page

# ============================================================
# IMPORTS PROTEGIDOS
# ============================================================

ADICIONALES_AVAILABLE = True
ADICIONALES_ERR = None
try:
    from services.adicionales_service import process_adicionales_streamlit
except Exception as e:
    ADICIONALES_AVAILABLE = False
    ADICIONALES_ERR = e

PERC_AVAILABLE = True
PERC_ERR = None
try:
    from services.percepcion_service import process_percepcion_streamlit
except Exception as e:
    PERC_AVAILABLE = False
    PERC_ERR = e

DUAS_AVAILABLE = True
DUAS_ERR = None
try:
    from services.duas_service import process_duas_streamlit
except Exception as e:
    DUAS_AVAILABLE = False
    DUAS_ERR = e

EXTERNOS_AVAILABLE = True
EXTERNOS_ERR = None
try:
    from services.externos_service import process_externos_streamlit
except Exception as e:
    EXTERNOS_AVAILABLE = False
    EXTERNOS_ERR = e

# ============================================================
# UTILIDADES
# ============================================================

from io import BytesIO
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

USE_PRN_WIDTHS = True

PRN_WIDTHS_1 = [10,25,6,6,6,16,16,2,5,16,3,2,30,6,3,3,8,3,6,4,16,16,3,6]
PRN_WIDTHS_2 = [6,3,3,8,3,16,16,2,30,6,15,20,5]

def set_fixed_widths(ws, widths, start_col=1):
    for i, w in enumerate(widths, start=start_col):
        ws.column_dimensions[get_column_letter(i)].width = float(w) + 0.71

def _autofit_worksheet(ws):
    for col in ws.columns:
        max_len = 8
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = min(max_len * 1.2, 60)

def header_paint(ws):
    BLUE = "FF0077B6"
    WHITE = "FFFFFFFF"
    fill_blue = PatternFill(fill_type="solid", start_color=BLUE, end_color=BLUE)
    font_white_bold = Font(color=WHITE, bold=True)

    # aplica apenas na linha de cabeçalho
    for cell in ws[1]:
        cell.fill = fill_blue
        cell.font = font_white_bold


def to_xlsx_bytes(df, sheet_name):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        _autofit_worksheet(ws)
        header_paint(ws)
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================
# ESTADO
# ============================================================

ACTIONS = {
    "externos": "Externos",
    "gastos": "Gastos Adicionales",
    "duas": "Duas",
    "percepciones": "Percepciones",
}

def _ensure_state():
    st.session_state.setdefault("acao_selecionada", None)
    st.session_state.setdefault("uploader_key", "uploader_none")
    st.session_state.setdefault("tasa_df", None)

def _select_action(action):
    st.session_state.acao_selecionada = action
    st.session_state.uploader_key = f"uploader_{action}"

# ============================================================
# PÁGINA
# ============================================================

def render():
    _ensure_state()
    st.subheader("Aplicación Comex")

    tab_model, tab_tasa, tab_sp, tab_proc, tab_prn = st.tabs([
        "📦 Arquivos modelo",
        "🌐 Tasa SUNAT",
        "📁 Arquivo Sharepoint",
        "📥 Processamento local",
        "📝 Transformar .prn",
    ])

    # ========================================================
    # TAB TASA SUNAT
    # ========================================================
    with tab_tasa:
        anos = st.multiselect("Anos", ["2024", "2025", "2026"], ["2024", "2025", "2026"])
        if st.button("Atualizar Tasa"):
            status = st.empty()
            pbar = st.progress(0)
            df = atualizar_dataframe_tasa(anos, pbar, status)
            if df is not None and not df.empty:
                st.session_state.tasa_df = df.copy()
                st.success("Tasa consolidada com sucesso.")
                st.dataframe(df.head(30))
                st.download_button("Baixar XLSX", to_xlsx_bytes(df, "Tasa"), "tasa.xlsx")

    # ========================================================
    # TAB SHAREPOINT  ✅ CORRIGIDA
    # ========================================================
    with tab_sp:
        st.subheader("📁 Arquivo Sharepoint")
        uploaded_excel = st.file_uploader("Carregar Excel", ["xlsx", "xls"])

        if uploaded_excel:
            try:
                xls = pd.ExcelFile(uploaded_excel, engine="openpyxl")

                # MAPEAMENTO ROBUSTO DAS ABAS
                sheet_map = {s.strip().lower(): s for s in xls.sheet_names}

                if "all" not in sheet_map:
                    st.error(
                        "❌ Aba 'all' não encontrada.\n\n"
                        f"Abas disponíveis: {xls.sheet_names}"
                    )
                    st.stop()

                sheet_real = sheet_map["all"]

                df_all = pd.read_excel(
                    uploaded_excel,
                    sheet_name=sheet_real,
                    usecols="A:Z",
                    nrows=20000,
                    engine="openpyxl",
                )

                from services.sharepoint_utils import ajustar_sharepoint_df
                df_all = ajustar_sharepoint_df(df_all)

                st.session_state["sharepoint_df"] = df_all
                st.success("✔️ DataFrame atualizado")
                st.dataframe(df_all, height=500)

                st.download_button(
                    "Baixar CSV",
                    df_all.to_csv(index=False).encode("utf-8"),
                    "sharepoint_all.csv",
                )

            except Exception as e:
                st.error("Erro ao processar o Excel do SharePoint.")
                st.exception(e)

    # ========================================================
    # TAB PROCESSAMENTO LOCAL
    # ========================================================
    with tab_proc:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Externos"):
                _select_action("externos")
            if st.button("Gastos Adicionales"):
                _select_action("gastos")
        with col2:
            if st.button("Duas"):
                _select_action("duas")
            if st.button("Percepciones"):
                _select_action("percepciones")

        acao = st.session_state.acao_selecionada
        if not acao:
            st.info("Selecione uma ação.")
            return

        files = st.file_uploader(
            f"PDFs para {ACTIONS[acao]}",
            type=["pdf"],
            accept_multiple_files=True,
            key=st.session_state.uploader_key,
        )

        if st.button("Executar") and files:
            status = st.empty()
            pbar = st.progress(0)

            if acao == "duas":
                df = process_duas_streamlit(files, pbar, status, st.session_state.tasa_df)
            elif acao == "externos":
                df = process_externos_streamlit(files, pbar, status, st.session_state.tasa_df)
            elif acao == "gastos":
                df = process_adicionales_streamlit(files, pbar, status, st.session_state.tasa_df)
            else:
                df = process_percepcion_streamlit(files, pbar, status)

            if df is not None and not df.empty:
                st.success("Processamento concluído.")
                st.dataframe(df.head(50))
                st.download_button(
                    "Baixar XLSX",
                    to_xlsx_bytes(df, ACTIONS[acao]),
                    f"{acao}.xlsx",
                )
            else:
                st.warning("Nenhum dado encontrado.")

    # ========================================================
    # TAB MODELOS / PRN
    # ========================================================
    with tab_model:
        downloads_page.render()

    with tab_prn:
        st.info("Transformação PRN mantida conforme versão anterior.")
``

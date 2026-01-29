# ui/pages/process_pdfs.py
import streamlit as st
from services import pdf_service  # (mantido mesmo que não utilizado agora)
from services.tasa_service import atualizar_dataframe_tasa

# --- Import protegido do Adicionales ---
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

# --- Import protegido do Percepciones ---
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

# --- Import protegido do DUAS ---
DUAS_AVAILABLE = True
DUAS_ERR = None
try:
    from services.duas_service import process_duas_streamlit
except Exception as e:
    DUAS_AVAILABLE = False
    DUAS_ERR = e

if not DUAS_AVAILABLE:
    st.warning(
        "O módulo **DUAS** não pôde ser carregado. "
        "Verifique `services/duas_service.py` e dependências (ex.: `pdfplumber`)."
    )
    with st.expander("Detalhes técnicos do erro (DUAS)"):
        st.exception(DUAS_ERR)  # <-- mostra o stack-trace real

# --- Import protegido do Externos ---
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

# -------------------
# Utilidades
# -------------------
from io import BytesIO
import pandas as pd


def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Dados") -> bytes:
    """
    Exporta DataFrame para XLSX com:
      - cabeçalho colorido e em negrito,
      - autofit (largura auto-ajustada por conteúdo),
      - filtro no cabeçalho,
      - primeira linha congelada.
    """
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # escreve cabeçalho + linhas
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # estilos do cabeçalho
    header_fill = PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid")  # laranja suave
    header_font = Font(bold=True, color="000000")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="D9D9D9")
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_col = ws.max_column
    max_row = ws.max_row

    # aplica estilo ao cabeçalho
    for col_idx in range(1, max_col + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = header_alignment
        c.border = header_border

    # autofit de largura
    base = 1.2  # multiplicador por caractere
    buffer = 3  # folga extra
    for col_idx in range(1, max_col + 1):
        letter = get_column_letter(col_idx)
        max_len = 0

        # cabeçalho
        hv = ws.cell(row=1, column=col_idx).value
        if hv is not None:
            max_len = max(max_len, len(str(hv)))

        # dados
        for row_idx in range(2, max_row + 1):
            v = ws.cell(row=row_idx, column=col_idx).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))

        min_width = 8
        ws.column_dimensions[letter].width = max(min_width, int(max_len * base) + buffer)

    # filtro e freeze
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    # salva bytes
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


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

    # Se houver chave inválida gravada, limpa para prevenir KeyError
    val = (st.session_state.acao_selecionada or "")
    if isinstance(val, str) and val.strip() not in ACTIONS:
        st.session_state.acao_selecionada = None
        st.session_state.uploader_key = "uploader_none"


def _select_action(action_key: str):
    key = (action_key or "").strip()  # normaliza (evita "externos " etc)
    st.session_state.acao_selecionada = key
    st.session_state.uploader_key = f"uploader_{key}"


def render():
    _ensure_state()

    st.subheader("Processar PDFs")
    tab1, tab2, tab3 = st.tabs(
        [
            "📥 Processamento local",
            "🌐 Tasa SUNAT",
            "📁 Arquivo Sharepoint",
        ]
    )

    # -------------------------------
    # 📥 Processamento local
    # -------------------------------
    with tab1:
        st.markdown("#### Ações rápidas")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Externos", use_container_width=True):
                _select_action("externos")
            if st.button("Gastos Adicionales", use_container_width=True):
                _select_action("gastos")
        with col2:
            if st.button("Duas", use_container_width=True):
                _select_action("duas")
            if st.button("Percepciones", use_container_width=True):
                _select_action("percepciones")

        # AVISO se DUAS não carregou (mantém botões visíveis)
        if not DUAS_AVAILABLE:
            st.warning(
                "O módulo **DUAS** não pôde ser carregado. "
                "Verifique `services/duas_service.py` e dependências (ex.: `pdfplumber`)."
            )

        # validação da ação selecionada
        acao_key = (st.session_state.get("acao_selecionada") or "").strip()
        has_action = bool(acao_key) and acao_key in ACTIONS

        if has_action:
            nome_acao = ACTIONS.get(acao_key, acao_key)
            st.info(f"🔧 Fluxo **{nome_acao}** selecionado.")
        else:
            # reseta se inválido/não selecionado
            st.session_state.acao_selecionada = None
            st.session_state.uploader_key = "uploader_none"
            st.caption("Selecione uma ação acima para enviar PDFs e executar o fluxo correspondente.")

        st.divider()

        # Uploader só aparece quando a ação é válida
        if has_action:
            uploaded_files = st.file_uploader(
                f"Envie um ou mais arquivos PDF para **{ACTIONS[acao_key]}**",
                type=["pdf"],
                accept_multiple_files=True,
                key=st.session_state.uploader_key,
                help="Os arquivos enviados serão processados pelo fluxo selecionado.",
            )
            col_run, col_clear = st.columns([2, 1])
            with col_run:
                run_clicked = st.button(
                    "▶️ Executar",
                    type="primary",
                    use_container_width=True,
                    disabled=not uploaded_files,
                )
            with col_clear:
                clear_clicked = st.button("Limpar seleção", use_container_width=True)

            if clear_clicked:
                st.session_state.acao_selecionada = None
                st.session_state.uploader_key = "uploader_none"
                st.rerun()

            # Execução do fluxo
            if run_clicked and uploaded_files:
                acao = (st.session_state.get("acao_selecionada") or "").strip()
                if acao not in ACTIONS:
                    st.error("Ação inválida ou não selecionada. Por favor, escolha um fluxo nos botões acima.")
                    st.stop()

                nome_acao = ACTIONS.get(acao, acao)
                status = st.empty()
                progress = st.progress(0, text=f"Iniciando fluxo {nome_acao}...")

                # ---------------- DUAS ----------------
                if acao == "duas":
                    cambio_df = st.session_state.get("tasa_df")
                    if cambio_df is None or getattr(cambio_df, "empty", True):
                        st.warning(
                            "Para calcular **Tasa**, primeiro atualize no tab **🌐 Tasa SUNAT**. "
                            "O processamento seguirá sem Tasa."
                        )
                    if not DUAS_AVAILABLE:
                        st.error("DUAS indisponível: confira dependências e arquivo `services/duas_service.py`.")
                    else:
                        df_final = process_duas_streamlit(
                            uploaded_files=uploaded_files,
                            progress_widget=progress,
                            status_widget=status,
                            cambio_df=cambio_df,
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
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="DUAS")
                                st.download_button(
                                    label="Baixar XLSX (DUAS)",
                                    data=xlsx_bytes,
                                    file_name="duas_consolidado.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                )
                        else:
                            st.warning("Nenhuma tabela válida encontrada nos PDFs para o fluxo DUAS.")

                # ------------- PERCEPCIONES -------------
                elif acao == "percepciones":
                    if not PERC_AVAILABLE:
                        st.error(
                            "Percepciones indisponível: confira dependências e `services/percepcion_service.py`."
                        )
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
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Percepciones")
                                st.download_button(
                                    label="Baixar XLSX (Percepciones)",
                                    data=xlsx_bytes,
                                    file_name="percepciones.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                )
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Percepciones.")

                # --------------- EXTERNOS ----------------
                elif acao == "externos":
                    if not EXTERNOS_AVAILABLE:
                        st.error("Externos indisponível: confira dependências e `services/externos_service.py`.")
                    else:
                        cambio_df = st.session_state.get("tasa_df")  # opcional
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
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Externos")
                                st.download_button(
                                    label="Baixar XLSX (Externos)",
                                    data=xlsx_bytes,
                                    file_name="externos.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                )
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Externos.")

                # ------------- GASTOS ADICIONALES -------------
                elif acao == "gastos":
                    if not ADICIONALES_AVAILABLE:
                        st.error(
                            "Gastos Adicionales indisponível: confira dependências e `services/adicionales_service.py`."
                        )
                    else:
                        cambio_df = st.session_state.get("tasa_df")  # opcional
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
                                )
                            with col_xlsx:
                                xlsx_bytes = to_xlsx_bytes(df_final, sheet_name="Adicionales")
                                st.download_button(
                                    label="Baixar XLSX (Adicionales)",
                                    data=xlsx_bytes,
                                    file_name="gastos_adicionales.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                )
                        else:
                            st.warning("Nenhuma informação válida encontrada nos PDFs para Gastos Adicionales.")

    # -------------------------------
    # 🌐 Tasa SUNAT
    # -------------------------------
    with tab2:
        st.write("Baixar e consolidar Tasa (SUNAT) direto do site oficial.")
        anos = st.multiselect(
            "Anos",
            ["2024", "2025", "2026"],
            default=["2024", "2025", "2026"],
        )
        if st.button("Atualizar Tasa"):
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
                    )
                with col_xlsx:
                    xlsx_bytes = to_xlsx_bytes(df, sheet_name="Tasa")
                    st.download_button(
                        label="Baixar XLSX",
                        data=xlsx_bytes,
                        file_name="tasa_consolidada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            else:
                st.warning("Não foi possível obter dados da Tasa. Verifique credenciais/token/cookie.")

    # -------------------------------
    # 📁 Arquivo Sharepoint
    # -------------------------------
    with tab3:
        st.subheader("📁 Arquivo Sharepoint")
        st.caption("Carregue um arquivo Excel para leitura da aba 'all'.")
        uploaded_excel = st.file_uploader(
            "Carregar Arquivo",
            type=["xlsx", "xls"],
            key="sharepoint_excel_uploader",
        )

        if uploaded_excel:
            try:
                df_all = pd.read_excel(
                    uploaded_excel,
                    sheet_name="all",
                    header=0,
                    usecols="A:Z",
                    nrows=20000,
                    engine="openpyxl",
                )
                from services.sharepoint_utils import ajustar_sharepoint_df

                df_all = ajustar_sharepoint_df(df_all)
                st.session_state["sharepoint_df"] = df_all
                st.success("✔️ DataFrame atualizado")
                st.dataframe(df_all, use_container_width=True, height=500)

                # Downloads do SharePoint
                st.subheader("⬇️ Downloads do Arquivo SharePoint")
                col_csv, col_xlsx = st.columns(2)
                with col_csv:
                    st.download_button(
                        label="Baixar CSV (SharePoint)",
                        data=df_all.to_csv(index=False).encode("utf-8"),
                        file_name="sharepoint_all.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )
                with col_xlsx:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df_all.to_excel(writer, index=False, sheet_name="SharePoint")
                    buffer.seek(0)
                    st.download_button(
                        label="Baixar XLSX (SharePoint)",
                        data=buffer,
                        file_name="sharepoint_all.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            except ValueError:
                st.error("❌ A aba 'all' não foi encontrada no arquivo Excel.")
            except Exception as e:
                st.error("❌ Erro ao processar o arquivo Excel.")
                st.exception(e)

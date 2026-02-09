# ui/pages/app_archivo_gastos.py
import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------------------------------------------
# Estado e helpers
# ------------------------------------------------------------
def _ensure_state():
    if "aag_state" not in st.session_state:
        st.session_state["aag_state"] = {
            "uploader_key": "aag_uploader_1",
            "last_action": None,
        }

def _set_action(action: str):
    st.session_state["aag_state"]["last_action"] = action

# ------------------------------------------------------------
# Página
# ------------------------------------------------------------
def render():
    _ensure_state()

    # Título da página
    st.subheader("Aplicación Archivo Gastos")

    # Descrição breve
    st.caption(
        "Espaço para processar/validar o **Arquivo de Gastos**. "
        "Nas próximas etapas, conectaremos as regras de negócio e exportações."
    )

    # --------------------------------------------------------
    # Seção de parâmetros (ex.: filtros, ano/mês, options etc.)
    # --------------------------------------------------------
    with st.expander("Parâmetros (opcional)", expanded=False):
        colp1, colp2, colp3 = st.columns(3)
        with colp1:
            ano = st.selectbox("Ano", ["2024", "2025", "2026"], index=1)
        with colp2:
            mes = st.selectbox("Mês", list(range(1, 13)), index=0)
        with colp3:
            modo = st.radio("Modo de execução", ["Validação", "Consolidação"], index=0)

    st.divider()

    # --------------------------------------------------------
    # Uploader (se a página for trabalhar com arquivos locais)
    # --------------------------------------------------------
    uploaded = st.file_uploader(
        "Carregar arquivo(s) de gastos",
        type=["xlsx", "xls", "csv", "txt", "pdf"],
        accept_multiple_files=True,
        key=st.session_state["aag_state"]["uploader_key"],
        help="Envie um ou mais arquivos conforme o fluxo do Arquivo de Gastos."
    )

    # Botões de ação
    col_run, col_clear = st.columns([2, 1])
    with col_run:
        run_clicked = st.button("▶️ Executar", type="primary", use_container_width=True, disabled=not uploaded)
    with col_clear:
        clear_clicked = st.button("Limpar", use_container_width=True)

    if clear_clicked:
        # Limpa seleção e reseta key do uploader
        st.session_state["aag_state"]["last_action"] = None
        st.session_state["aag_state"]["uploader_key"] = st.session_state["aag_state"]["uploader_key"] + "_x"
        st.rerun()

    # --------------------------------------------------------
    # Resultado (placeholder) - aqui entra seu pipeline real
    # --------------------------------------------------------
    if run_clicked and uploaded:
        status = st.empty()
        pbar = st.progress(0, text="Iniciando processamento...")

        try:
            # Exemplo mínimo: apenas lista nomes
            nomes = [getattr(f, "name", "arquivo") for f in uploaded]
            df_preview = pd.DataFrame({"Arquivos recebidos": nomes})
            pbar.progress(50, text="Lendo estrutura...")

            # TODO: Conectar seu pipeline real aqui
            # TODO: Normalizações, merges, cálculos, export...

            pbar.progress(100, text="Concluído.")
            st.success("Processamento finalizado com sucesso.")
            st.dataframe(df_preview, use_container_width=True)

            # Exemplo de export simples para XLSX
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_preview.to_excel(writer, index=False, sheet_name="Preview")
            buffer.seek(0)
            st.download_button(
                "Baixar XLSX (preview)",
                data=buffer.getvalue(),
                file_name="archivo_gastos_preview.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error("Erro durante o processamento.")
            st.exception(e)
    else:
        st.info("Envie arquivo(s) e clique em **Executar** para iniciar.")

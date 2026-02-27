# app.py
import streamlit as st
from auth import is_authenticated
from ui.login import render_login
from ui.layout import app_header, sidebar_navigation
from settings import PAGES, APP_NAME
from ui.pages import home, process_pdfs, settings_page
from ui.pages import downloads_page
from ui.pages import app_archivo_gastos

try:
    # Quando o app roda como package (Streamlit Cloud / execução padrão)
    from .asientos_contables_module import render_asientos_contables_ui
except ImportError:
    # Fallback local (se executar via `streamlit run comex_pdf_reader/app.py`)
    from asientos_contables_module import render_asientos_contables_ui
def main():
    st.set_page_config(page_title="COMEX PDF READER", page_icon="📄", layout="wide")

    if not is_authenticated():
        render_login()
        return

    # 1) Ler a página escolhida (ANTES do header)
    page = sidebar_navigation(PAGES)

    # 2) Se houve navegação por botão, priorizar esse destino
    if st.session_state.get("_goto_page"):
        page = st.session_state.pop("_goto_page")

    # 3) Definir o título dinâmico do cabeçalho
    header_title = APP_NAME  # padrão
    if page == "Aplicación Archivo Gastos":
        header_title = "PLANTILLA GASTOS"  # título em espanhol, como você pediu

    # 4) Renderizar o cabeçalho com o título escolhido
    app_header(title=header_title)
    
    # 5) Roteamento
    if page == "Home":
        home.render()
    elif page == "Aplicación Comex":
        process_pdfs.render()
    elif page == "Aplicación Archivo Gastos":
        app_archivo_gastos.render()
    elif page == "Asientos Contables":                      # ← NOVO
        render_asientos_contables_ui(session_key_df="asientos_df")
    elif page == "Configurações":
        settings_page.render()
    else:
        st.error("Página não encontrada.")
        
if __name__ == "__main__":
    main()

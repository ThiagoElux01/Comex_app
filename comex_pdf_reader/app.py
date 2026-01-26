import streamlit as st
from auth import is_authenticated
from ui.login import render_login
from ui.layout import app_header, sidebar_navigation
from settings import PAGES
from ui.pages import home, process_pdfs, settings_page
from ui.pages import downloads_page

def main():
    st.set_page_config(page_title="COMEX PDF READER", page_icon="📄", layout="wide")

    if not is_authenticated():
        render_login()
        return

    app_header()

    # 1) lê a página escolhida
    page = sidebar_navigation(PAGES)

    # 3) se foi disparada navegação por botão, priorize esse destino
    if st.session_state.get("_goto_page"):
        page = st.session_state.pop("_goto_page")

    # 4) roteamento
    if page == "Home":
        home.render()
    elif page == "Processar PDFs":
        process_pdfs.render()
    elif page == "Arquivos modelo":
        downloads_page.render()
    elif page == "Configurações":
        settings_page.render()
    else:
        st.error("Página não encontrada.")

if __name__ == "__main__":
    main()

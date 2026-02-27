# app.py
import streamlit as st
from auth import is_authenticated
from ui.login import render_login
from ui.layout import app_header, sidebar_navigation
from settings import PAGES, APP_NAME
from ui.pages import home, process_pdfs, settings_page
from ui.pages import downloads_page
from ui.pages import app_archivo_gastos

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
        header_title = "Plantilla Gastos"  # título em espanhol, como você pediu

    # 4) Renderizar o cabeçalho com o título escolhido
    app_header(title=header_title)

    # 5) Roteamento
    if page == "Home":
        home.render()
    elif page == "Aplicación Comex":
        process_pdfs.render()
    elif page == "Aplicación Archivo Gastos":
        app_archivo_gastos.render()
    elif page == "Configurações":
        settings_page.render()
    else:
        st.error("Página não encontrada.")

if __name__ == "__main__":
    main()

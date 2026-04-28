# app.py
import streamlit as st

from auth import is_authenticated
from ui.login import render_login
from ui.layout import app_header, sidebar_navigation
from settings import PAGES, APP_NAME

from ui.pages import (
    home,
    process_pdfs,
    settings_page,
    downloads_page,
    app_archivo_gastos,
)


def main():
    # ----------------- Configuração da página -----------------
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="📄",
        layout="wide",
    )

    # ----------------- Login -----------------
    if not is_authenticated():
        render_login()
        return

    # ----------------- Navegação -----------------
    page = sidebar_navigation(PAGES)

    # Override por navegação programática (botões, fluxos guiados)
    if "_goto_page" in st.session_state:
        page = st.session_state.pop("_goto_page")

    # ----------------- Header dinâmico -----------------
    header_title = APP_NAME
    if page == "Aplicación Archivo Gastos":
        header_title = "PLANTILLA GASTOS"

    app_header(title=header_title)

    # ----------------- Roteamento -----------------
    if page == "Home":
        home.render()

    elif page == "Aplicación Comex":
        process_pdfs.render()

    elif page == "Aplicación Archivo Gastos":
        app_archivo_gastos.render()

    elif page == "Configurações":
        settings_page.render()

    elif page == "Downloads":
        downloads_page.render()

    else:
        st.error(f"Página não encontrada: {page}")


# ----------------- Bootstrap -----------------
if __name__ == "__main__":
    main()

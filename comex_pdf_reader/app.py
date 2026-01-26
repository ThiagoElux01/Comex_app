import streamlit as st
from auth import is_authenticated
from ui.login import render_login
from ui.layout import app_header, sidebar_navigation
from settings import PAGES
from ui.pages import home, process_pdfs, settings_page

def main():
    st.set_page_config(page_title="COMEX PDF READER", page_icon="📄", layout="wide")

    if not is_authenticated():
        render_login()
        return

    app_header()
    page = sidebar_navigation(PAGES)

    if page == "Home":
        home.render()
    elif page == "Processar PDFs":
        process_pdfs.render()
    elif page == "Configurações":
        settings_page.render()
    else:
        st.error("Página não encontrada.")

if __name__ == "__main__":
    main()

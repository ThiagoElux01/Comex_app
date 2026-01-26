import streamlit as st
from settings import APP_NAME
from auth import do_logout

def app_header():
    left, mid, right = st.columns([1, 2, 1])
    with left:
        st.markdown(f"### 📄 {APP_NAME}")
    with mid:
        st.empty()
    with right:
        st.caption(f"Usuário: **{st.session_state.get('user_email', '')}**")
        st.button("Sair", on_click=do_logout)
    st.divider()

def sidebar_navigation(pages: list[str]) -> str:
    with st.sidebar:
        st.header("Menu")
        return st.radio("Navegação", pages, index=0)


import streamlit as st
from settings import APP_NAME
# ⚠️ você pode manter o import do do_logout se ainda usar em outro lugar,
# mas aqui vamos fazer o logout sem callback.

def app_header():
    left, mid, right = st.columns([1, 2, 1])
    with left:
        st.markdown(f"### 📄 {APP_NAME}")
    with mid:
        st.empty()
    with right:
        st.caption(f"Usuário: **{st.session_state.get('user_email', '')}**")

        # ✅ Sem callback: trata o clique aqui no fluxo principal
        if st.button("Sair"):
            # Limpa o estado de autenticação
            st.session_state["is_logged_in"] = False
            st.session_state["user_email"] = ""
            # Se houver mais chaves de sessão de login, limpe-as aqui também:
            # for k in ("roles", "token", "nome", ...): st.session_state.pop(k, None)

            # ✅ Agora sim, pode forçar o rerun fora de callback
            st.rerun()

    st.divider()

def sidebar_navigation(pages: list[str]) -> str:
    with st.sidebar:
        st.header("Menu")
        return st.radio("Navegação", pages, index=0)

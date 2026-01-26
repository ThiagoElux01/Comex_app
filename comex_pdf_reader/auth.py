
import streamlit as st

def _get_authorized_users() -> dict:
    """Lê usuários autorizados de st.secrets['auth']."""
    return dict(st.secrets.get("auth", {}))

def login(email: str, password: str) -> bool:
    """Valida credenciais usando st.secrets."""
    if not email:
        return False
    expected = _get_authorized_users().get(email.strip())
    return expected is not None and password == expected

def is_authenticated() -> bool:
    return st.session_state.get("authenticated", False) is True

def set_authenticated(email: str):
    st.session_state["authenticated"] = True
    st.session_state["user_email"] = email

def do_logout():
    for k in ["authenticated", "user_email"]:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

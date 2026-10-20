import bcrypt
import streamlit as st


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(password: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))
    except Exception:
        return False


def login_user(user: dict) -> None:
    st.session_state.user = user


def logout_user() -> None:
    st.session_state.pop("user", None)
    st.session_state.pop("cart", None)


def current_user():
    return st.session_state.get("user")


def has_role(*roles: str) -> bool:
    user = current_user()
    return bool(user and user.get("role") in roles)

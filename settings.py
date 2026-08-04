"""
Read configuration from wherever the host provides it.

Different hosts pass settings in different ways:

- Streamlit Community Cloud -> st.secrets (the "Secrets" box in the app settings)
- Azure App Service / Docker / local -> environment variables

Checking both means the same code runs unchanged on any of them.
"""

import os

import streamlit as st


def get_setting(name: str, default: str = "") -> str:
    """Value for `name` from st.secrets, falling back to the environment."""
    try:
        if name in st.secrets:
            return str(st.secrets[name]).strip()
    except Exception:
        # No secrets file configured - normal when running from env vars.
        pass
    return (os.environ.get(name) or default).strip()

"""
Access gate for the names pipeline app.

This app handles beneficiary names, phone numbers and national ID numbers, so it
must not be reachable without some form of access control.

Two supported setups:

1. **Platform SSO (preferred).** Azure App Service Authentication (Entra ID)
   sits in front of the app, so requests never reach Python unauthenticated.
   Set AUTH_MODE=platform to tell this module the gate is handled upstream.

2. **Shared password (fallback).** When SSO isn't available - e.g. the tenant
   blocks app registrations - set APP_PASSWORD and the app prompts for it.
   Weaker than SSO: no per-person audit trail, and revoking access for one
   person means rotating the password for everyone.

Fail-secure by design: if neither is configured the app refuses to start rather
than silently serving data to anyone who finds the URL. ALLOW_NO_AUTH=true
overrides that, and exists only for local development.
"""

import hmac
import os
import time

import streamlit as st

_ATTEMPT_KEY = "_auth_attempts"
_AUTHED_KEY = "_auth_ok"


def _platform_user() -> str | None:
    """Username injected by Azure App Service Authentication, if present.

    App Service adds these headers after a successful Entra ID sign-in. They
    cannot be spoofed by the client because App Service strips inbound copies.
    """
    try:
        headers = st.context.headers or {}
    except Exception:
        return None
    for header in ("X-MS-CLIENT-PRINCIPAL-NAME", "X-MS-CLIENT-PRINCIPAL-ID"):
        value = headers.get(header) or headers.get(header.lower())
        if value:
            return value
    return None


def require_access() -> str:
    """Gate the app. Returns a label for who's signed in, or stops the script."""
    mode = (os.environ.get("AUTH_MODE") or "").strip().lower()
    password = os.environ.get("APP_PASSWORD") or ""
    allow_open = (os.environ.get("ALLOW_NO_AUTH") or "").strip().lower() == "true"

    # ── 1. Platform SSO in front of the app ────────────────────────────────
    if mode == "platform":
        user = _platform_user()
        if user:
            return user
        # Configured for SSO but no principal header arrived. Don't fall back to
        # serving the app - that would defeat the gate.
        st.error(
            "AUTH_MODE is set to `platform`, but no signed-in user was found.\n\n"
            "App Service Authentication may not be switched on yet. Until it is, "
            "this app will not open.",
            icon="🔒",
        )
        st.stop()

    # ── 2. Shared password ────────────────────────────────────────────────
    if password:
        if st.session_state.get(_AUTHED_KEY):
            return "shared password"

        st.title("Names Pipeline")
        st.caption("Enter the team password to continue.")
        with st.form("login"):
            entered = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in", type="primary")

        if submitted:
            # Constant-time compare so response timing doesn't leak the password.
            if hmac.compare_digest(entered, password):
                st.session_state[_AUTHED_KEY] = True
                st.rerun()
            else:
                attempts = st.session_state.get(_ATTEMPT_KEY, 0) + 1
                st.session_state[_ATTEMPT_KEY] = attempts
                # Small escalating delay to make guessing impractical.
                time.sleep(min(attempts, 5))
                st.error("Incorrect password.", icon="🚫")
        st.stop()

    # ── 3. Nothing configured ─────────────────────────────────────────────
    if allow_open:
        st.warning(
            "**Running with no access control** (`ALLOW_NO_AUTH=true`). Fine on "
            "your own machine — never for a hosted URL, because this app handles "
            "beneficiary names, phone numbers and national IDs.",
            icon="⚠️",
        )
        return "no auth (local)"

    st.error(
        "**This app is not configured for access control, so it will not start.**\n\n"
        "It handles beneficiary personal data, so it must not be reachable "
        "without a sign-in. Set one of the following:\n\n"
        "- `AUTH_MODE=platform` — Azure App Service Authentication (Entra ID) is "
        "in front of the app *(preferred)*\n"
        "- `APP_PASSWORD=<password>` — shared team password\n"
        "- `ALLOW_NO_AUTH=true` — **local development only**",
        icon="🔒",
    )
    st.stop()


def sign_out_button() -> None:
    """Sign-out control, shown only for the shared-password mode.

    Platform SSO sign-out is handled by App Service at /.auth/logout.
    """
    if st.session_state.get(_AUTHED_KEY):
        if st.sidebar.button("Sign out"):
            st.session_state.pop(_AUTHED_KEY, None)
            st.rerun()

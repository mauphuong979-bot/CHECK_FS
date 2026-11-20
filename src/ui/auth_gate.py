"""Màn hình đăng nhập, ghi nhớ đăng nhập, menu tài khoản và đăng xuất."""

from __future__ import annotations

import time

import streamlit as st

from src.auth import usage_log
from src.auth.models import UserAccount
from src.auth.remember_token import create_token, read_username, verify_token
from src.auth.settings import AuthSettings, load_auth_settings
from src.auth.user_store import UserDirectory, get_user_directory
from src.ui import browser_cookies
from src.ui.branding import APP_ICON, APP_TITLE

MAX_FAILED_ATTEMPTS = 5
LOCKOUT_SECONDS = 30
_AUTH_USER = "auth_user"
_AUTH_TOKEN_VERSION = "auth_token_version"
_PENDING_COOKIE_SCRIPT = "auth_cookie_script"
_PENDING_NOTICE = "auth_notice"
_COVER_MAIN_PAGE = "auth_cover_main_page"
_LOGIN_FORM_SHOWN = "auth_login_form_shown"
ENTRY_OVERLAY_GIVE_UP_SECONDS = 15
ENTRY_OVERLAY_HTML = f"""<style>
.auth-wait {{ position: fixed; inset: 0; z-index: 999990; background: #FFFFFF; display: flex;
    flex-direction: column; align-items: center; justify-content: center; gap: 14px;
    font-family: "Inter", "Segoe UI", sans-serif; color: #61758A; font-size: 15px;
    animation: auth-give-up 0s {ENTRY_OVERLAY_GIVE_UP_SECONDS}s forwards; }}
.auth-wait-spinner {{ width: 38px; height: 38px; border-radius: 50%; border: 4px solid #E6E6E6;
    border-top-color: #1F4E78; animation: auth-rotate 0.8s linear infinite; }}
.auth-wait-fallback {{ font-size: 0.85rem; visibility: hidden; animation: auth-reveal 0s 8s forwards; }}
@keyframes auth-rotate {{ to {{ transform: rotate(360deg); }} }}
@keyframes auth-reveal {{ to {{ visibility: visible; }} }}
@keyframes auth-give-up {{ to {{ visibility: hidden; }} }}
</style>
<div class="auth-wait">
    <div class="auth-wait-spinner"></div>
    <div>Đang kiểm tra đăng nhập…</div>
    <div class="auth-wait-fallback">Chưa tự đăng nhập được.
        <a href="?login=password" target="_self">Đăng nhập bằng mật khẩu</a></div>
</div>"""
_SESSION_DATA_KEYS_TO_CLEAR_ON_SIGN_OUT = (
    "tables",
    "workbook_bytes",
    "commented_docx_bytes",
    "attention_items",
    "regime_detection",
    "applied_rule_pack",
    "coverage",
    "processed_filename",
    "processed_file_fingerprint",
)


def open_entry_overlay():
    """Lớp phủ "Đang kiểm tra đăng nhập…"; phải là phần tử đầu tiên của trang.

    Được vẽ cùng nội dung, cùng vị trí ở mọi lượt chạy cho đến khi form đăng nhập hoặc toàn bộ trang chính
    đã vẽ xong, nên trình duyệt giữ nguyên một màn hình chờ thay vì để lộ trang đang dựng dở giữa các lần
    chạy lại. Lớp phủ tự ẩn sau ``ENTRY_OVERLAY_GIVE_UP_SECONDS`` giây để không che lỗi bất ngờ.
    """
    state = st.session_state
    if state.get(_COVER_MAIN_PAGE) or (not state.get(_AUTH_USER) and not state.get(_LOGIN_FORM_SHOWN)):
        return st.html(ENTRY_OVERLAY_HTML)
    return st.empty()


def close_entry_overlay(overlay) -> None:
    """Gỡ lớp phủ sau khi trang chính đã vẽ xong."""
    st.session_state.pop(_COVER_MAIN_PAGE, None)
    overlay.empty()


def require_login(overlay) -> tuple[UserAccount, UserDirectory]:
    """Tài khoản đang đăng nhập; nếu chưa đăng nhập thì vẽ màn hình đăng nhập và dừng lượt chạy."""
    settings = load_auth_settings()
    directory = get_user_directory(settings)
    state = st.session_state

    if _PENDING_COOKIE_SCRIPT in state:
        _run_script(state.pop(_PENDING_COOKIE_SCRIPT))

    current = _current_user(directory)
    if current is not None:
        return current, directory

    if settings.remember_enabled and _try_remembered_sign_in(settings, directory):
        st.rerun()

    if settings.remember_enabled and _should_ask_browser_for_token():
        _render_bridge_wait_screen()
        st.stop()

    _render_login_screen(settings, directory)
    st.session_state[_LOGIN_FORM_SHOWN] = True
    overlay.empty()  # chỉ gỡ khi form đã vẽ xong, không có khoảng trống nhấp nháy
    st.stop()
    raise RuntimeError("unreachable")  # st.stop() luôn dừng lượt chạy.


def render_account_menu(user: UserAccount, directory: UserDirectory) -> None:
    with st.popover(f"👤 {user.username}", help="Tài khoản và đăng xuất"):
        st.markdown(f"**{user.username}**  \n{user.role_label}")
        with st.expander("🔑 Đổi mật khẩu"):
            _render_change_password_form(user, directory)
        if st.button("🚪 Đăng xuất", key="sign_out", width="stretch"):
            sign_out(user)


def sign_out(user: UserAccount, *, notice: str | None = None) -> None:
    usage_log.log_event(user.username, usage_log.EVENT_LOGOUT)
    state = st.session_state
    for key in (_AUTH_USER, _AUTH_TOKEN_VERSION, *_SESSION_DATA_KEYS_TO_CLEAR_ON_SIGN_OUT):
        state.pop(key, None)
    state[_PENDING_COOKIE_SCRIPT] = browser_cookies.clear_cookie_script()
    state["auth_signed_out"] = True
    if notice:
        state[_PENDING_NOTICE] = notice
    st.rerun()


def refresh_session_version(directory: UserDirectory, username: str) -> None:
    """Giữ phiên hiện tại sau khi chính người dùng đổi mật khẩu (token_version đã tăng)."""
    user = directory.find(username)
    if user is not None:
        st.session_state[_AUTH_TOKEN_VERSION] = user.token_version


def _current_user(directory: UserDirectory) -> UserAccount | None:
    username = st.session_state.get(_AUTH_USER)
    if not username:
        return None
    user = directory.find(username)
    if user is None:
        sign_out(UserAccount(username, ""), notice="Tài khoản không còn tồn tại. Vui lòng liên hệ quản trị viên.")
    session_version = st.session_state.get(_AUTH_TOKEN_VERSION, user.token_version)
    if directory.warning is None and user.token_version != session_version:
        sign_out(user, notice="Phiên đăng nhập đã hết hiệu lực do mật khẩu hoặc quyền truy cập thay đổi. "
                              "Vui lòng đăng nhập lại.")
    return user


def _sign_in(user: UserAccount) -> None:
    state = st.session_state
    state[_AUTH_USER] = user.username
    state[_AUTH_TOKEN_VERSION] = user.token_version
    state.pop("auth_signed_out", None)
    state.pop("auth_failed_attempts", None)
    state[_COVER_MAIN_PAGE] = True  # giữ màn hình chờ đến khi trang chính vẽ xong


def _try_remembered_sign_in(settings: AuthSettings, directory: UserDirectory) -> bool:
    state = st.session_state
    if state.get("auth_signed_out"):
        return False
    bridge_value = state.get(browser_cookies.BRIDGE_INPUT_KEY)
    if bridge_value == browser_cookies.NO_TOKEN or st.query_params.get("login") == "password":
        state["auth_bridge_done"] = True
        bridge_value = None
    token = st.context.cookies.get(browser_cookies.COOKIE_NAME) or bridge_value
    if not token or token == state.get("auth_tried_token"):
        return False
    state["auth_tried_token"] = token
    state["auth_bridge_done"] = True

    user = directory.find(read_username(token) or "")
    if user is not None and verify_token(
        token,
        secret=settings.cookie_secret,
        password_hash=user.password_hash,
        token_version=user.token_version,
    ):
        _sign_in(user)
        return True
    state[_PENDING_COOKIE_SCRIPT] = browser_cookies.clear_cookie_script()
    return False


def _should_ask_browser_for_token() -> bool:
    state = st.session_state
    return not (
        st.context.cookies.get(browser_cookies.COOKIE_NAME)
        or state.get("auth_tried_token")
        or state.get("auth_bridge_done")
        or state.get("auth_signed_out")
    )


def _render_bridge_wait_screen() -> None:
    """Ô nhập ẩn và script cầu nối; lớp phủ chờ đã do ``open_entry_overlay()`` vẽ."""
    with st.container(key=f"{browser_cookies.BRIDGE_INPUT_KEY}_box"):
        st.text_input("remember", key=browser_cookies.BRIDGE_INPUT_KEY, label_visibility="collapsed")
    _run_script(browser_cookies.bridge_script())


def _render_login_screen(settings: AuthSettings, directory: UserDirectory) -> None:
    _, column, _ = st.columns([1, 1.15, 1])
    with column, st.container(border=True, key="login_box"):
        st.markdown(
            f'<div class="login-logo">{APP_ICON}</div>'
            f'<div class="login-header">{APP_TITLE}</div>'
            '<div class="login-sub">🔐 Đăng nhập để tiếp tục</div>',
            unsafe_allow_html=True,
        )
        notice = st.session_state.pop(_PENDING_NOTICE, None)
        if notice:
            st.info(notice)

        users = directory.users()
        if directory.warning:
            st.warning(directory.warning)
        if not users:
            st.error(
                "Chưa có tài khoản nào được cấu hình. Quản trị viên cần khai báo tài khoản trong mục "
                "[auth.users] của Streamlit Secrets hoặc trong Google Sheet."
            )
            return

        with st.form("login_form"):
            username = st.selectbox(
                "Tên đăng nhập",
                [user.username for user in users],
                index=None,
                placeholder="Chọn tên đăng nhập",
            )
            password = st.text_input("Mật khẩu", type="password")
            remember = st.checkbox(
                "Ghi nhớ đăng nhập trên máy này",
                value=settings.remember_enabled,
                disabled=not settings.remember_enabled,
                help=(
                    f"Không cần nhập mật khẩu trong {settings.remember_days} ngày trên trình duyệt này. "
                    "Không dùng trên máy dùng chung."
                    if settings.remember_enabled
                    else "Quản trị viên chưa cấu hình khóa ghi nhớ đăng nhập (auth.cookie_secret)."
                ),
            )
            submitted = st.form_submit_button("Đăng nhập", type="primary", width="stretch")

        if submitted:
            _handle_login(directory, settings, username, password, remember)
        st.caption("Quên mật khẩu? Liên hệ quản trị viên để được cấp lại.")


def _handle_login(
    directory: UserDirectory,
    settings: AuthSettings,
    username: str | None,
    password: str,
    remember: bool,
) -> None:
    locked_seconds = _remaining_lockout_seconds()
    if locked_seconds:
        st.error(f"Nhập sai quá nhiều lần. Vui lòng thử lại sau {locked_seconds} giây.")
        return
    if not username:
        st.error("Vui lòng chọn tên đăng nhập.")
        return
    if not password:
        st.error("Vui lòng nhập mật khẩu.")
        return
    user = directory.authenticate(username, password)
    if user is None:
        _record_failed_attempt()
        st.error("Sai tên đăng nhập hoặc mật khẩu.")
        return
    _sign_in(user)
    if remember and settings.remember_enabled:
        token = create_token(
            user.username,
            secret=settings.cookie_secret,
            password_hash=user.password_hash,
            token_version=user.token_version,
            lifetime_seconds=settings.remember_seconds,
        )
        st.session_state[_PENDING_COOKIE_SCRIPT] = browser_cookies.set_cookie_script(
            token, settings.remember_seconds
        )
    st.rerun()


def _render_change_password_form(user: UserAccount, directory: UserDirectory) -> None:
    if not directory.writable:
        st.caption("Chưa cấu hình Google Sheet nên chưa thể đổi mật khẩu. Vui lòng liên hệ quản trị viên.")
        return
    with st.form("change_password_form", clear_on_submit=True):
        current_password = st.text_input("Mật khẩu hiện tại", type="password")
        new_password = st.text_input("Mật khẩu mới", type="password")
        confirm_password = st.text_input("Nhập lại mật khẩu mới", type="password")
        submitted = st.form_submit_button("Lưu mật khẩu mới", width="stretch")
    if not submitted:
        return
    if not current_password or not new_password:
        st.warning("Vui lòng nhập mật khẩu hiện tại và mật khẩu mới.")
    elif new_password != confirm_password:
        st.error("Hai lần nhập mật khẩu mới không khớp.")
    else:
        result = directory.change_password(user.username, current_password, new_password)
        if result.ok:
            refresh_session_version(directory, user.username)
            usage_log.log_event(user.username, usage_log.EVENT_ACCOUNT, "Tự đổi mật khẩu")
            st.success(result.message)
        else:
            st.error(result.message)


def _record_failed_attempt() -> None:
    state = st.session_state
    attempts = state.get("auth_failed_attempts", 0) + 1
    state["auth_failed_attempts"] = attempts
    if attempts >= MAX_FAILED_ATTEMPTS:
        state["auth_locked_until"] = time.time() + LOCKOUT_SECONDS
        state["auth_failed_attempts"] = 0


def _remaining_lockout_seconds() -> int:
    remaining = st.session_state.get("auth_locked_until", 0) - time.time()
    return int(remaining) + 1 if remaining > 0 else 0


def _run_script(script: str) -> None:
    st.components.v1.html(script, height=0)

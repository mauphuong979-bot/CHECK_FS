"""Tab "Quản trị": trạng thái kết nối, quản lý tài khoản và nhật ký sử dụng (chỉ Quản trị viên)."""

from __future__ import annotations

import csv
from datetime import datetime
import io

import streamlit as st

from src.auth import gsheet, usage_log
from src.auth.models import ROLE_LABELS, ROLES, ROLE_USER, UserAccount
from src.auth.user_store import StoreResult, UserDirectory
from src.ui.auth_gate import refresh_session_version

_FLASH_KEY = "admin_flash"


def render_admin_page(current_user: UserAccount, directory: UserDirectory) -> None:
    _render_flash()
    _render_connection_status(directory)
    with st.container(border=True, key="admin_users_box"):
        st.markdown('<div class="section-title">👤 Quản lý tài khoản</div>', unsafe_allow_html=True)
        _render_add_user_form(current_user, directory)
        _render_user_rows(current_user, directory)
    with st.container(border=True, key="admin_logs_box"):
        st.markdown('<div class="section-title">📈 Nhật ký sử dụng</div>', unsafe_allow_html=True)
        _render_usage_logs()


def _render_connection_status(directory: UserDirectory) -> None:
    if not gsheet.is_configured():
        st.warning(
            "Chưa cấu hình Google Sheet. Đang dùng tài khoản khai báo trong Streamlit Secrets ở chế độ chỉ đọc: "
            "chưa thể thêm/sửa/xóa tài khoản, đổi mật khẩu hay lưu nhật ký sử dụng. "
            "Bổ sung mục [connections.gsheets] trong Secrets để bật các chức năng này."
        )
    elif directory.warning:
        st.warning(directory.warning)
    else:
        st.caption("🟢 Đã kết nối Google Sheet: tài khoản lưu tại worksheet **Users**, nhật ký tại **Logs**.")


def _render_add_user_form(current_user: UserAccount, directory: UserDirectory) -> None:
    with st.expander("➕ Thêm tài khoản", expanded=False):
        with st.form("add_user_form", clear_on_submit=True):
            name_col, password_col, role_col = st.columns([1.2, 1.2, 1])
            username = name_col.text_input("Tên đăng nhập", placeholder="Ví dụ: VPHN")
            password = password_col.text_input("Mật khẩu", type="password", help="Tối thiểu 6 ký tự.")
            role = role_col.selectbox(
                "Vai trò", ROLES, index=ROLES.index(ROLE_USER), format_func=ROLE_LABELS.get
            )
            submitted = st.form_submit_button(
                "➕ Thêm tài khoản", type="primary", disabled=not directory.writable
            )
        if submitted:
            if not username.strip() or not password:
                st.warning("Vui lòng nhập cả tên đăng nhập và mật khẩu.")
            else:
                _apply(
                    directory.add_user(username, password, role),
                    current_user,
                    f"Thêm tài khoản {username.strip()} ({role})",
                    directory=directory,
                )


def _render_user_rows(current_user: UserAccount, directory: UserDirectory) -> None:
    users = directory.users()
    if not users:
        st.info("Chưa có tài khoản nào. Dùng mục **➕ Thêm tài khoản** ở trên để tạo tài khoản đầu tiên.")
        return

    st.caption("Để trống ô **Mật khẩu mới** nếu không đổi mật khẩu. Đổi tên hoặc mật khẩu sẽ buộc đăng nhập lại.")
    header = st.columns([1.3, 1.3, 1.1, 0.6, 0.6, 0.6], vertical_alignment="bottom")
    for column, label in zip(header, ("Tên đăng nhập", "Mật khẩu mới", "Vai trò", "", "", "")):
        column.markdown(f'<div class="admin-col-head">{label}</div>', unsafe_allow_html=True)

    disabled = not directory.writable
    for user in users:
        index = user.username
        is_self = user.username == current_user.username
        columns = st.columns([1.3, 1.3, 1.1, 0.6, 0.6, 0.6], vertical_alignment="bottom")
        new_name = columns[0].text_input(
            "Tên đăng nhập", value=user.username, key=f"admin_name_{index}",
            label_visibility="collapsed", disabled=disabled,
        )
        new_password = columns[1].text_input(
            "Mật khẩu mới", type="password", key=f"admin_password_{index}", placeholder="Giữ nguyên",
            label_visibility="collapsed", disabled=disabled,
        )
        new_role = columns[2].selectbox(
            "Vai trò", ROLES, index=ROLES.index(user.role), format_func=ROLE_LABELS.get,
            key=f"admin_role_{index}", label_visibility="collapsed", disabled=disabled,
        )
        if columns[3].button("💾", key=f"admin_save_{index}", help=f"Lưu thay đổi của {user.username}",
                             width="stretch", disabled=disabled):
            _apply(
                directory.update_user(user.username, new_username=new_name, new_password=new_password, role=new_role),
                current_user,
                f"Cập nhật tài khoản {user.username}",
                directory=directory,
                self_username=new_name.strip() if is_self else None,
            )
        if columns[4].button("🚪", key=f"admin_revoke_{index}",
                             help=f"Hủy ghi nhớ đăng nhập của {user.username} trên mọi trình duyệt",
                             width="stretch", disabled=disabled or is_self):
            _apply(
                directory.revoke_remembered(user.username),
                current_user,
                f"Hủy ghi nhớ đăng nhập của {user.username}",
                directory=directory,
            )
        if columns[5].button("🗑️", key=f"admin_delete_{index}",
                             help="Không thể xóa tài khoản đang đăng nhập" if is_self else f"Xóa {user.username}",
                             width="stretch", disabled=disabled or is_self):
            _confirm_delete(user.username, current_user, directory)


def _confirm_delete(username: str, current_user: UserAccount, directory: UserDirectory) -> None:
    def render() -> None:
        st.warning(f"Xóa tài khoản **{username}**? Thao tác này không thể hoàn tác.")
        confirm_col, cancel_col = st.columns(2)
        if confirm_col.button("🗑️ Xóa tài khoản", type="primary", width="stretch", key="admin_confirm_delete"):
            _apply(
                directory.delete_user(username, acting_username=current_user.username),
                current_user,
                f"Xóa tài khoản {username}",
                directory=directory,
            )
        if cancel_col.button("Hủy", width="stretch", key="admin_cancel_delete"):
            st.rerun()

    st.dialog("Xác nhận xóa tài khoản")(render)()


def _apply(
    result: StoreResult,
    current_user: UserAccount,
    log_details: str,
    *,
    directory: UserDirectory,
    self_username: str | None = None,
) -> None:
    """Hiển thị kết quả sau lượt chạy kế tiếp; giữ phiên khi quản trị viên tự sửa tài khoản của mình."""
    state = st.session_state
    if not result.ok:
        state[_FLASH_KEY] = ("error", result.message)
        st.rerun()
    usage_log.log_event(current_user.username, usage_log.EVENT_ACCOUNT, log_details)
    state[_FLASH_KEY] = ("success", result.message)
    for key in [key for key in state if str(key).startswith("admin_password_")]:
        del state[key]
    if self_username:
        refresh_session_version(directory, self_username)
        state["auth_user"] = self_username
    st.rerun()


def _render_flash() -> None:
    flash = st.session_state.pop(_FLASH_KEY, None)
    if flash:
        kind, message = flash
        (st.success if kind == "success" else st.error)(message)


def _render_usage_logs() -> None:
    if not gsheet.is_configured():
        st.info("Nhật ký sử dụng sẽ hiển thị tại đây sau khi cấu hình Google Sheet.")
        return
    try:
        events = usage_log.read_events()
    except gsheet.SheetUnavailableError:
        st.error("Không đọc được nhật ký từ Google Sheet. Vui lòng thử lại sau.")
        return
    if not events:
        st.info("Chưa có nhật ký sử dụng.")
        return

    user_col, event_col, search_col = st.columns([1, 1, 1.3])
    selected_users = user_col.multiselect(
        "Người dùng", sorted({event["Người dùng"] for event in events}), placeholder="Tất cả"
    )
    selected_events = event_col.multiselect(
        "Sự kiện", sorted({event["Sự kiện"] for event in events}), placeholder="Tất cả"
    )
    keyword = search_col.text_input("Tìm trong chi tiết", placeholder="Nhập từ khóa")
    filtered = [
        event for event in events
        if (not selected_users or event["Người dùng"] in selected_users)
        and (not selected_events or event["Sự kiện"] in selected_events)
        and (not keyword or keyword.lower() in event["Chi tiết"].lower())
    ]
    st.dataframe(filtered, width="stretch", hide_index=True, height=min(420, 36 * (len(filtered) + 1) + 4))

    count_col, download_col = st.columns([2, 1], vertical_alignment="center")
    count_col.caption(f"Đang hiển thị {len(filtered)}/{len(events)} bản ghi.")
    download_col.download_button(
        "📥 Tải nhật ký đã lọc (CSV)",
        data=_events_to_csv(filtered),
        file_name=f"nhat_ky_su_dung_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        width="stretch",
        disabled=not filtered,
    )


def _events_to_csv(events: list[dict[str, str]]) -> bytes:
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=usage_log.LOGS_HEADERS)
    writer.writeheader()
    writer.writerows(events)
    return buffer.getvalue().encode("utf-8-sig")

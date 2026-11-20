"""Kết nối Google Sheet dùng chung cho tài khoản và nhật ký sử dụng.

Cấu hình đặt tại ``[connections.gsheets]`` trong Streamlit Secrets: khóa service account cùng URL bảng tính
(``spreadsheet``). Khi chưa cấu hình, ứng dụng vẫn chạy với tài khoản khai báo trong ``[auth.users]``.
"""

from __future__ import annotations

from collections.abc import Callable
import re

import streamlit as st

from src.auth.settings import read_secret_section

_SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
)


class SheetUnavailableError(RuntimeError):
    """Không kết nối được Google Sheet; thông điệp an toàn để hiển thị cho người dùng."""


def is_configured() -> bool:
    config = read_secret_section("connections", "gsheets")
    return bool(_spreadsheet_url(config) and config.get("private_key") and config.get("client_email"))


def open_worksheet(
    title: str,
    headers: list[str],
    *,
    seed_rows: Callable[[], list[list[str]]] | None = None,
):
    """Worksheet theo tên; tự tạo kèm dòng tiêu đề (và dữ liệu khởi tạo) nếu chưa có."""
    import gspread

    spreadsheet = _open_spreadsheet()
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        rows = [headers, *(seed_rows() if seed_rows else [])]
        worksheet = spreadsheet.add_worksheet(title=title, rows=max(200, len(rows) + 50), cols=len(headers))
        worksheet.update(values=rows, range_name="A1", value_input_option="RAW")
        return worksheet


def _open_spreadsheet():
    if not is_configured():
        raise SheetUnavailableError("Chưa cấu hình Google Sheet trong Secrets.")
    try:
        return _cached_spreadsheet()
    except Exception as exc:
        raise SheetUnavailableError("Không kết nối được Google Sheet.") from exc


@st.cache_resource(show_spinner=False)
def _cached_spreadsheet():
    import gspread
    from google.oauth2.service_account import Credentials

    config = dict(read_secret_section("connections", "gsheets"))
    url = _spreadsheet_url(config)
    for key in ("spreadsheet", "spreadsheet_url"):
        config.pop(key, None)
    config["private_key"] = normalize_private_key(str(config.get("private_key", "")))
    credentials = Credentials.from_service_account_info(config, scopes=list(_SCOPES))
    return gspread.authorize(credentials).open_by_url(url)


def _spreadsheet_url(config) -> str:
    return str(config.get("spreadsheet") or config.get("spreadsheet_url") or "")


def normalize_private_key(key: str) -> str:
    """Dựng lại khóa PEM bị dán sai dạng: ``\\n`` dạng chữ, xuống dòng Windows, thụt lề hoặc nằm trên một dòng."""
    text = key.replace("\\\\n", "\n").replace("\\n", "\n").replace("\r", "")
    match = re.search(r"-----BEGIN ([A-Z ]+)-----(.*?)-----END \1-----", text, re.S)
    if not match:
        return text
    body = re.sub(r"[^A-Za-z0-9+/=]", "", match.group(2))
    lines = [body[index:index + 64] for index in range(0, len(body), 64)]
    return "\n".join([f"-----BEGIN {match.group(1)}-----", *lines, f"-----END {match.group(1)}-----"]) + "\n"

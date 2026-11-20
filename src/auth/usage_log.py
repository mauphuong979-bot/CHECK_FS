"""Nhật ký sử dụng ghi vào worksheet ``Logs`` của Google Sheet.

Ghi người dùng, sự kiện và chi tiết vận hành; với lần kiểm tra BCTC có thêm tên tệp Word để truy vết, nhưng không
ghi nội dung báo cáo. Bảng tính nhật ký chỉ nên chia sẻ cho quản trị viên. Log máy chủ chỉ ghi loại sự kiện và
người dùng (không ghi chi tiết); lỗi Google Sheet không làm gián đoạn người dùng.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
import logging

from src.auth import gsheet

LOGGER = logging.getLogger(__name__)
LOGS_SHEET_TITLE = "Logs"
LOGS_HEADERS = ["Thời gian", "Người dùng", "Sự kiện", "Chi tiết"]
VN_TIMEZONE = timezone(timedelta(hours=7))

EVENT_LOGOUT = "Đăng xuất"
EVENT_AUDIT = "Kiểm tra BCTC"
EVENT_AUDIT_FAILED = "Kiểm tra BCTC lỗi"
EVENT_ACCOUNT = "Quản lý tài khoản"


def log_event(username: str, event: str, details: str = "") -> bool:
    """Ghi một sự kiện; trả ``True`` khi đã lưu lên Google Sheet."""
    LOGGER.info("Sự kiện %s của %s.", event, username)
    if not gsheet.is_configured():
        return False
    row = [datetime.now(VN_TIMEZONE).strftime("%Y-%m-%d %H:%M:%S"), username, event, details]
    try:
        gsheet.open_worksheet(LOGS_SHEET_TITLE, LOGS_HEADERS).append_row(row, value_input_option="RAW")
        return True
    except Exception:
        LOGGER.warning("Không ghi được nhật ký sử dụng lên Google Sheet.")
        return False


def read_events() -> list[dict[str, str]]:
    """Toàn bộ nhật ký, mới nhất trước. Ném ``SheetUnavailableError`` khi không đọc được."""
    try:
        rows = gsheet.open_worksheet(LOGS_SHEET_TITLE, LOGS_HEADERS).get_all_values()[1:]
    except gsheet.SheetUnavailableError:
        raise
    except Exception as exc:
        raise gsheet.SheetUnavailableError("Không đọc được nhật ký sử dụng.") from exc
    events = [
        dict(zip(LOGS_HEADERS, (list(row) + [""] * len(LOGS_HEADERS))[: len(LOGS_HEADERS)]))
        for row in rows
        if any(cell.strip() for cell in row)
    ]
    return events[::-1]

"""Danh bạ tài khoản: đọc/ghi worksheet ``Users`` của Google Sheet, dự phòng bằng ``[auth.users]`` trong Secrets.

- Google Sheet là nguồn chính, dùng chung giữa máy cục bộ và Streamlit Cloud; lần đầu worksheet được tạo và
  khởi tạo từ danh sách tài khoản trong Secrets (mật khẩu được băm trước khi ghi).
- Khi chưa cấu hình hoặc không kết nối được Google Sheet, đăng nhập dùng danh sách trong Secrets ở chế độ chỉ
  đọc; mọi thao tác thay đổi tài khoản trả lỗi rõ ràng thay vì âm thầm bỏ qua.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
import logging
from typing import Protocol

import streamlit as st

from src.auth import gsheet
from src.auth.models import ROLE_ADMIN, UserAccount, normalize_role, username_problem
from src.auth.passwords import hash_password, password_problem, verify_password
from src.auth.settings import AuthSettings

LOGGER = logging.getLogger(__name__)
USERS_SHEET_TITLE = "Users"
USERS_HEADERS = ["username", "password_hash", "role", "token_version"]
USERS_CACHE_SECONDS = 60


@dataclass(frozen=True)
class StoreResult:
    ok: bool
    message: str


class UserBackend(Protocol):
    writable: bool

    def read(self, *, use_cache: bool = True) -> list[UserAccount]: ...

    def write(self, users: list[UserAccount]) -> None: ...


class SecretsUserBackend:
    """Danh sách tài khoản khai báo trong Secrets; chỉ đọc."""

    writable = False

    def __init__(self, users: tuple[UserAccount, ...]):
        self._users = list(users)

    def read(self, *, use_cache: bool = True) -> list[UserAccount]:
        return list(self._users)

    def write(self, users: list[UserAccount]) -> None:
        raise gsheet.SheetUnavailableError("Chưa cấu hình Google Sheet trong Secrets.")


class SheetUserBackend:
    writable = True

    def __init__(self, seed_users: tuple[UserAccount, ...]):
        self._seed_users = seed_users

    def read(self, *, use_cache: bool = True) -> list[UserAccount]:
        rows = _cached_sheet_rows(self._seed_users) if use_cache else _read_sheet_rows(self._seed_users)
        return [user for user in (_row_to_user(row) for row in rows) if user is not None]

    def write(self, users: list[UserAccount]) -> None:
        try:
            worksheet = gsheet.open_worksheet(USERS_SHEET_TITLE, USERS_HEADERS)
            previous_count = len(worksheet.get_all_values())
            rows = [USERS_HEADERS, *(_user_to_row(user) for user in users)]
            if len(rows) > worksheet.row_count:
                worksheet.add_rows(len(rows) - worksheet.row_count + 50)
            worksheet.update(values=rows, range_name="A1", value_input_option="RAW")
            if previous_count > len(rows):
                worksheet.batch_clear([f"A{len(rows) + 1}:D{previous_count}"])
        except gsheet.SheetUnavailableError:
            raise
        except Exception as exc:
            raise gsheet.SheetUnavailableError("Lỗi khi lưu lên Google Sheet.") from exc
        finally:
            _cached_sheet_rows.clear()


class UserDirectory:
    def __init__(
        self,
        backend: UserBackend,
        *,
        fallback_users: tuple[UserAccount, ...] = (),
    ):
        self._backend = backend
        self._fallback_users = fallback_users
        self.warning: str | None = None

    @property
    def writable(self) -> bool:
        return self._backend.writable

    def users(self) -> list[UserAccount]:
        try:
            users = self._backend.read()
            self.warning = None
            return users
        except gsheet.SheetUnavailableError:
            LOGGER.warning("Không đọc được danh sách tài khoản từ Google Sheet; dùng danh sách dự phòng.")
            self.warning = (
                "Không kết nối được Google Sheet, đang dùng danh sách tài khoản dự phòng trong Secrets. "
                "Mật khẩu đã đổi gần đây có thể chưa có hiệu lực."
            )
            return list(self._fallback_users)

    def find(self, username: str) -> UserAccount | None:
        return next((user for user in self.users() if user.username == username), None)

    def authenticate(self, username: str, password: str) -> UserAccount | None:
        user = self.find(username)
        return user if user is not None and verify_password(password, user.password_hash) else None

    def add_user(self, username: str, password: str, role: str) -> StoreResult:
        username = username.strip()
        problem = username_problem(username) or password_problem(password)
        if problem:
            return StoreResult(False, problem)

        def modify(users: list[UserAccount]) -> str | None:
            if any(user.username.lower() == username.lower() for user in users):
                return f"Tên đăng nhập '{username}' đã tồn tại."
            users.append(UserAccount(username, hash_password(password), normalize_role(role)))
            return None

        return self._change(modify, f"Đã thêm tài khoản '{username}'.")

    def update_user(
        self,
        username: str,
        *,
        new_username: str,
        new_password: str = "",
        role: str,
    ) -> StoreResult:
        new_username = new_username.strip()
        problem = username_problem(new_username) or (password_problem(new_password) if new_password else None)
        if problem:
            return StoreResult(False, problem)

        def modify(users: list[UserAccount]) -> str | None:
            index = _index_of(users, username)
            if index is None:
                return f"Không tìm thấy tài khoản '{username}'."
            if new_username.lower() != username.lower() and any(
                user.username.lower() == new_username.lower() for user in users
            ):
                return f"Tên đăng nhập '{new_username}' đã tồn tại."
            current = users[index]
            if current.is_admin and normalize_role(role) != ROLE_ADMIN and _admin_count(users) == 1:
                return "Cần giữ ít nhất một tài khoản Quản trị viên."
            users[index] = replace(
                current,
                username=new_username,
                role=normalize_role(role),
                password_hash=hash_password(new_password) if new_password else current.password_hash,
                token_version=current.token_version + (1 if new_password or new_username != username else 0),
            )
            return None

        return self._change(modify, f"Đã cập nhật tài khoản '{new_username}'.")

    def delete_user(self, username: str, *, acting_username: str) -> StoreResult:
        if username == acting_username:
            return StoreResult(False, "Không thể xóa tài khoản đang đăng nhập.")

        def modify(users: list[UserAccount]) -> str | None:
            index = _index_of(users, username)
            if index is None:
                return f"Không tìm thấy tài khoản '{username}'."
            if users[index].is_admin and _admin_count(users) == 1:
                return "Cần giữ ít nhất một tài khoản Quản trị viên."
            del users[index]
            return None

        return self._change(modify, f"Đã xóa tài khoản '{username}'.")

    def change_password(self, username: str, current_password: str, new_password: str) -> StoreResult:
        problem = password_problem(new_password)
        if problem:
            return StoreResult(False, problem)
        if new_password == current_password:
            return StoreResult(False, "Mật khẩu mới phải khác mật khẩu hiện tại.")

        def modify(users: list[UserAccount]) -> str | None:
            index = _index_of(users, username)
            if index is None:
                return "Không tìm thấy tài khoản."
            if not verify_password(current_password, users[index].password_hash):
                return "Mật khẩu hiện tại không đúng."
            users[index] = replace(
                users[index],
                password_hash=hash_password(new_password),
                token_version=users[index].token_version + 1,
            )
            return None

        return self._change(
            modify,
            "Đã đổi mật khẩu. Các trình duyệt khác đã ghi nhớ đăng nhập sẽ phải đăng nhập lại.",
        )

    def revoke_remembered(self, username: str) -> StoreResult:
        def modify(users: list[UserAccount]) -> str | None:
            index = _index_of(users, username)
            if index is None:
                return f"Không tìm thấy tài khoản '{username}'."
            users[index] = replace(users[index], token_version=users[index].token_version + 1)
            return None

        return self._change(modify, f"Đã hủy ghi nhớ đăng nhập của '{username}' trên mọi trình duyệt.")

    def _change(self, modify: Callable[[list[UserAccount]], str | None], done_message: str) -> StoreResult:
        if not self._backend.writable:
            return StoreResult(
                False,
                "Chưa cấu hình Google Sheet nên chưa thể lưu thay đổi tài khoản. "
                "Vui lòng liên hệ quản trị viên để bổ sung cấu hình.",
            )
        try:
            users = self._backend.read(use_cache=False)
            error = modify(users)
            if error:
                return StoreResult(False, error)
            self._backend.write(users)
        except gsheet.SheetUnavailableError:
            LOGGER.warning("Không lưu được thay đổi tài khoản lên Google Sheet.")
            return StoreResult(False, "Không kết nối được Google Sheet, chưa lưu thay đổi. Vui lòng thử lại sau.")
        return StoreResult(True, done_message)


def get_user_directory(settings: AuthSettings) -> UserDirectory:
    if gsheet.is_configured():
        backend: UserBackend = SheetUserBackend(settings.bootstrap_users)
    else:
        backend = SecretsUserBackend(settings.bootstrap_users)
    return UserDirectory(backend, fallback_users=settings.bootstrap_users)


@st.cache_data(ttl=USERS_CACHE_SECONDS, show_spinner=False)
def _cached_sheet_rows(seed_users: tuple[UserAccount, ...]) -> list[list[str]]:
    return _read_sheet_rows(seed_users)


def _read_sheet_rows(seed_users: tuple[UserAccount, ...]) -> list[list[str]]:
    def seed_rows() -> list[list[str]]:
        return [
            _user_to_row(replace(user, password_hash=hash_password(user.password_hash)))
            for user in seed_users
        ]

    try:
        worksheet = gsheet.open_worksheet(USERS_SHEET_TITLE, USERS_HEADERS, seed_rows=seed_rows)
        return worksheet.get_all_values()[1:]
    except gsheet.SheetUnavailableError:
        raise
    except Exception as exc:
        raise gsheet.SheetUnavailableError("Không đọc được danh sách tài khoản.") from exc


def _row_to_user(row: list[str]) -> UserAccount | None:
    cells = (list(row) + [""] * len(USERS_HEADERS))[: len(USERS_HEADERS)]
    username = cells[0].strip()
    if not username or not cells[1]:
        return None
    try:
        token_version = int(cells[3] or 0)
    except ValueError:
        token_version = 0
    return UserAccount(username, cells[1], normalize_role(cells[2]), token_version)


def _user_to_row(user: UserAccount) -> list[str]:
    return [user.username, user.password_hash, user.role, str(user.token_version)]


def _index_of(users: list[UserAccount], username: str) -> int | None:
    return next((index for index, user in enumerate(users) if user.username == username), None)


def _admin_count(users: list[UserAccount]) -> int:
    return sum(1 for user in users if user.is_admin)

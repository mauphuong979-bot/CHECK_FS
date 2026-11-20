from __future__ import annotations

from dataclasses import dataclass
import re

ROLE_ADMIN = "admin"
ROLE_USER = "user"
ROLES = (ROLE_ADMIN, ROLE_USER)
ROLE_LABELS = {ROLE_ADMIN: "Quản trị viên", ROLE_USER: "Người dùng"}

_USERNAME_PATTERN = re.compile(r"^[A-Za-z0-9_.-]{2,32}$")


@dataclass(frozen=True)
class UserAccount:
    """Tài khoản đăng nhập. ``password_hash`` có thể là giá trị băm hoặc mật khẩu rõ nhập tay."""

    username: str
    password_hash: str
    role: str = ROLE_USER
    token_version: int = 0

    @property
    def is_admin(self) -> bool:
        return self.role == ROLE_ADMIN

    @property
    def role_label(self) -> str:
        return ROLE_LABELS.get(self.role, ROLE_LABELS[ROLE_USER])


def normalize_role(value: object) -> str:
    text = str(value or "").strip().lower()
    return text if text in ROLES else ROLE_USER


def username_problem(username: str) -> str | None:
    if not _USERNAME_PATTERN.fullmatch(username):
        return "Tên đăng nhập gồm 2–32 ký tự: chữ không dấu, số, dấu chấm, gạch ngang hoặc gạch dưới."
    return None

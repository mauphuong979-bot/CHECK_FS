"""Băm và xác minh mật khẩu bằng PBKDF2-SHA256 (thư viện chuẩn, không lưu mật khẩu dạng rõ)."""

from __future__ import annotations

import base64
import hashlib
import hmac
import secrets

HASH_PREFIX = "pbkdf2_sha256"
DEFAULT_ITERATIONS = 240_000
MIN_PASSWORD_LENGTH = 6


def hash_password(password: str, *, iterations: int = DEFAULT_ITERATIONS) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return "$".join(
        (HASH_PREFIX, str(iterations), _b64encode(salt), _b64encode(digest))
    )


def is_password_hash(value: str) -> bool:
    return value.startswith(f"{HASH_PREFIX}$")


def verify_password(password: str, stored: str) -> bool:
    """So khớp mật khẩu với giá trị đã lưu.

    Giá trị không theo định dạng băm được coi là mật khẩu rõ do quản trị viên nhập tay vào
    Google Sheet; vẫn so sánh theo thời gian hằng để tránh lộ thông tin qua thời gian phản hồi.
    """
    if not stored:
        return False
    if not is_password_hash(stored):
        return hmac.compare_digest(password.encode("utf-8"), stored.encode("utf-8"))
    try:
        _, iterations, salt, expected = stored.split("$")
        digest = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), _b64decode(salt), int(iterations)
        )
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(digest, _b64decode(expected))


def password_problem(password: str) -> str | None:
    """Thông báo lỗi thân thiện nếu mật khẩu mới không đạt yêu cầu tối thiểu."""
    if len(password) < MIN_PASSWORD_LENGTH:
        return f"Mật khẩu cần tối thiểu {MIN_PASSWORD_LENGTH} ký tự."
    if password.strip() != password:
        return "Mật khẩu không được bắt đầu hoặc kết thúc bằng khoảng trắng."
    return None


def _b64encode(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def _b64decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))

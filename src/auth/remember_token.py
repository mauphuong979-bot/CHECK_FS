"""Token "Ghi nhớ đăng nhập" ký bằng HMAC, lưu trong cookie trình duyệt.

Token chỉ chứa tên đăng nhập và thời hạn; chữ ký gắn với khóa bí mật của ứng dụng, giá trị băm mật khẩu
hiện tại và ``token_version`` của tài khoản. Vì vậy đổi mật khẩu hoặc quản trị viên tăng ``token_version``
("Hủy ghi nhớ") sẽ vô hiệu mọi token cũ mà không cần bảng lưu phiên phía máy chủ.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import time

_SEPARATOR = "|"


def create_token(
    username: str,
    *,
    secret: str,
    password_hash: str,
    token_version: int,
    lifetime_seconds: int,
    now: float | None = None,
) -> str:
    expires_at = int((now if now is not None else time.time()) + lifetime_seconds)
    payload = f"{username}{_SEPARATOR}{expires_at}"
    signature = _sign(payload, secret, password_hash, token_version)
    encoded = base64.urlsafe_b64encode(payload.encode("utf-8")).decode("ascii").rstrip("=")
    return f"{encoded}.{signature}"


def read_username(token: str | None) -> str | None:
    """Tên đăng nhập ghi trong token (chưa xác minh chữ ký), dùng để tra tài khoản."""
    payload = _payload(token)
    return payload[0] if payload else None


def verify_token(
    token: str | None,
    *,
    secret: str,
    password_hash: str,
    token_version: int,
    now: float | None = None,
) -> bool:
    payload = _payload(token)
    if payload is None or not secret:
        return False
    username, expires_at = payload
    if expires_at < (now if now is not None else time.time()):
        return False
    expected = _sign(
        f"{username}{_SEPARATOR}{expires_at}", secret, password_hash, token_version
    )
    return hmac.compare_digest(expected, token.rsplit(".", 1)[1])


def _payload(token: str | None) -> tuple[str, int] | None:
    if not token or token.count(".") != 1:
        return None
    encoded, _ = token.split(".")
    try:
        text = base64.urlsafe_b64decode(encoded + "=" * (-len(encoded) % 4)).decode("utf-8")
        username, expires_text = text.rsplit(_SEPARATOR, 1)
        return (username, int(expires_text)) if username else None
    except (ValueError, UnicodeDecodeError):
        return None


def _sign(payload: str, secret: str, password_hash: str, token_version: int) -> str:
    message = _SEPARATOR.join((payload, password_hash, str(token_version))).encode("utf-8")
    return hmac.new(secret.encode("utf-8"), message, hashlib.sha256).hexdigest()

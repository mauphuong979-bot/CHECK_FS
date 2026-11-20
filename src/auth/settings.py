"""Đọc cấu hình đăng nhập từ Streamlit Secrets (``.streamlit/secrets.toml`` hoặc Secrets trên Cloud)."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
import logging

import streamlit as st

from src.auth.models import UserAccount, normalize_role

LOGGER = logging.getLogger(__name__)
DEFAULT_REMEMBER_DAYS = 30


@dataclass(frozen=True)
class AuthSettings:
    cookie_secret: str = ""
    remember_days: int = DEFAULT_REMEMBER_DAYS
    bootstrap_users: tuple[UserAccount, ...] = field(default_factory=tuple)

    @property
    def remember_enabled(self) -> bool:
        return len(self.cookie_secret) >= 16

    @property
    def remember_seconds(self) -> int:
        return self.remember_days * 24 * 3600


def load_auth_settings() -> AuthSettings:
    section = read_secret_section("auth")
    return AuthSettings(
        cookie_secret=str(section.get("cookie_secret", "") or ""),
        remember_days=_positive_int(section.get("remember_days"), DEFAULT_REMEMBER_DAYS),
        bootstrap_users=tuple(_bootstrap_users(section.get("users"))),
    )


def read_secret_section(*path: str) -> Mapping:
    """Một nhánh của ``st.secrets`` dưới dạng dict; trả rỗng khi chưa có tệp hoặc chưa khai báo."""
    try:
        node: object = st.secrets
        for key in path:
            if not isinstance(node, Mapping) or key not in node:
                return {}
            node = node[key]
        return _to_plain(node) if isinstance(node, Mapping) else {}
    except Exception:  # Streamlit báo lỗi riêng khi không tìm thấy secrets.toml.
        LOGGER.info("Không đọc được Streamlit Secrets cho nhánh %s.", ".".join(path))
        return {}


def _bootstrap_users(raw: object) -> list[UserAccount]:
    if not isinstance(raw, Mapping):
        return []
    users = []
    for username, entry in raw.items():
        if not isinstance(entry, Mapping) or not str(entry.get("password", "")):
            continue
        users.append(
            UserAccount(
                username=str(username).strip(),
                password_hash=str(entry["password"]),
                role=normalize_role(entry.get("role")),
            )
        )
    return users


def _to_plain(node: Mapping) -> dict:
    return {
        key: _to_plain(value) if isinstance(value, Mapping) else value
        for key, value in node.items()
    }


def _positive_int(value: object, default: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return default
    return number if number > 0 else default

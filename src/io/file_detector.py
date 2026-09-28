from __future__ import annotations

from pathlib import Path


SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".csv", ".txt", ".doc", ".docm", ".xlsb"}


def detect_extension(filename: str) -> str:
    return Path(filename).suffix.lower()


def is_supported(filename: str) -> bool:
    return detect_extension(filename) in SUPPORTED_EXTENSIONS

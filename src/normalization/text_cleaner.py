from __future__ import annotations

import re
import unicodedata


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = unicodedata.normalize("NFC", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalized_key(value: object) -> str:
    text = clean_text(value).lower()
    text = "".join(
        char for char in unicodedata.normalize("NFD", text)
        if unicodedata.category(char) != "Mn"
    )
    text = text.replace("đ", "d")
    return re.sub(r"\s+", " ", text).strip()

from __future__ import annotations

import re

from src.domain.models import (
    AccountingRegime,
    AccountingRegimeDetection,
    DetectionConfidence,
    ExtractedTable,
)
from src.normalization.text_cleaner import clean_text, normalized_key


_EXPLICIT_PATTERNS: dict[AccountingRegime, re.Pattern[str]] = {
    AccountingRegime.TT200_2014: re.compile(r"(?:thông\s*tư|tt)\s*(?:số\s*)?200\s*/\s*2014\s*/\s*tt\s*-?\s*btc", re.IGNORECASE),
    AccountingRegime.TT133_2016: re.compile(r"(?:thông\s*tư|tt)\s*(?:số\s*)?133\s*/\s*2016\s*/\s*tt\s*-?\s*btc", re.IGNORECASE),
    AccountingRegime.TT99_2025: re.compile(r"(?:thông\s*tư|tt)\s*(?:số\s*)?99\s*/\s*2025\s*/\s*tt\s*-?\s*btc", re.IGNORECASE),
}


def detect_accounting_regime(
    tables: list[ExtractedTable],
    extra_fragments: tuple[str, ...] = (),
) -> AccountingRegimeDetection:
    fragments = (*extra_fragments, *_document_fragments(tables))
    matches: dict[AccountingRegime, str] = {}
    for fragment in fragments:
        for regime, pattern in _EXPLICIT_PATTERNS.items():
            if regime not in matches and pattern.search(fragment):
                text = clean_text(fragment)
                if len(text) > 180:
                    text = text[:180].rsplit(" ", 1)[0] + "..."
                matches[regime] = text

    if len(matches) == 1:
        regime, evidence = next(iter(matches.items()))
        return AccountingRegimeDetection(
            regime=regime,
            confidence=DetectionConfidence.HIGH,
            evidence=(evidence,),
            message="Nhận diện từ nội dung viện dẫn Thông tư trong báo cáo.",
        )
    if len(matches) > 1:
        names = ", ".join(f"Thông tư {regime.value}" for regime in matches)
        return AccountingRegimeDetection(
            regime=AccountingRegime.AMBIGUOUS,
            confidence=DetectionConfidence.LOW,
            evidence=tuple(matches.values()),
            message=f"Phát hiện nhiều văn bản có thể áp dụng: {names}. Cần người dùng xác minh.",
        )

    tt200_evidence = _detect_tt200_structure(tables)
    if tt200_evidence:
        return AccountingRegimeDetection(
            regime=AccountingRegime.TT200_2014,
            confidence=DetectionConfidence.MEDIUM,
            evidence=tt200_evidence,
            message="Suy luận từ ký hiệu biểu mẫu và cấu trúc mã chỉ tiêu đặc trưng; cần đối chiếu chính sách kế toán nếu có.",
        )

    tt133_evidence = _detect_tt133_structure(tables)
    if tt133_evidence:
        return AccountingRegimeDetection(
            regime=AccountingRegime.TT133_2016,
            confidence=DetectionConfidence.MEDIUM,
            evidence=tt133_evidence,
            message="Suy luận từ ký hiệu biểu mẫu và cấu trúc mã chỉ tiêu đặc trưng (TT 133/2016/TT-BTC); cần đối chiếu chính sách kế toán nếu có.",
        )

    return AccountingRegimeDetection(
        regime=AccountingRegime.UNKNOWN,
        confidence=DetectionConfidence.NONE,
        message="Không tìm thấy căn cứ đủ tin cậy để xác định Thông tư áp dụng.",
    )


def _document_fragments(tables: list[ExtractedTable]) -> tuple[str, ...]:
    fragments: list[str] = []
    for table in tables:
        fragments.extend((table.title_hint, *table.context_hints))
        fragments.extend(" ".join(row) for row in table.rows)
    return tuple(fragment for fragment in fragments if clean_text(fragment))


def _detect_tt200_structure(tables: list[ExtractedTable]) -> tuple[str, ...]:
    keys = [normalized_key(fragment) for fragment in _document_fragments(tables)]
    form_markers = tuple(
        marker
        for marker in ("b01-dn", "b02-dn", "b03-dn")
        if any(marker in key or marker.replace("-", " ") in key for key in keys)
    )
    if len(form_markers) >= 2:
        return (f"Phát hiện các ký hiệu biểu mẫu: {', '.join(form_markers).upper()}.",)
    return ()


def _detect_tt133_structure(tables: list[ExtractedTable]) -> tuple[str, ...]:
    keys = [normalized_key(fragment) for fragment in _document_fragments(tables)]
    form_markers = tuple(
        marker
        for marker in ("b01a-dnn", "b01b-dnn", "b02-dnn", "b03-dnn")
        if any(marker in key or marker.replace("-", " ") in key for key in keys)
    )
    if len(form_markers) >= 2:
        return (f"Phát hiện các ký hiệu biểu mẫu: {', '.join(form_markers).upper()}.",)
    return ()

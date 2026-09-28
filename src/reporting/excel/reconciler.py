from __future__ import annotations

from src.domain.models import ExtractedTable, NoteMatchResult


def _ordered_note_matches(matches: list[NoteMatchResult]) -> list[NoteMatchResult]:
    return sorted(
        matches,
        key=lambda match: (
            match.note_table_index,
            -match.match_score,
            match.statement_code,
        ),
    )

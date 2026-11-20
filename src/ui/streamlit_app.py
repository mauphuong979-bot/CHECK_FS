from __future__ import annotations

from hashlib import sha256
from html import escape
import logging
from pathlib import Path
from tempfile import NamedTemporaryFile
from time import perf_counter

import streamlit as st

from src.auth import usage_log
from src.auth.models import UserAccount
from src.auth.user_store import UserDirectory
from src.reporting.number_format import format_amount
from src.services.audit_service import (
    AuditAttentionItem,
    AuditCoverage,
    AuditOutputs,
    UnsupportedFileError,
    create_audit_outputs,
)
from src.ui.admin_page import render_admin_page
from src.ui.auth_gate import render_account_menu, require_login
from src.ui.branding import APP_ICON, APP_TITLE, TRANSLATION_COMPARE_URL
from src.ui.user_guide import show_user_guide

LOGGER = logging.getLogger(__name__)
STYLES_PATH = Path(__file__).with_name("styles.css")


def main() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon=APP_ICON,
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    _apply_styles()
    user, directory = require_login()
    _render_top_bar(user, directory)

    if user.is_admin:
        audit_tab, admin_tab = st.tabs(["🔍 Kiểm tra BCTC", "⚙️ Quản trị"])
        with audit_tab:
            render_audit_workspace(user.username)
        with admin_tab:
            render_admin_page(user, directory)
    else:
        render_audit_workspace(user.username)


def _render_top_bar(user: UserAccount, directory: UserDirectory) -> None:
    with st.container(key="top_bar", horizontal=True, vertical_alignment="center"):
        st.markdown(f'<div class="app-title">{APP_ICON} {APP_TITLE}</div>', unsafe_allow_html=True)
        st.space("stretch")
        st.link_button(
            "🌐 Đối chiếu số với bản dịch",
            TRANSLATION_COMPARE_URL,
            help="Mở ứng dụng so khớp số giữa BCTC tiếng Việt và bản dịch (tab mới).",
        )
        if st.button("❓", key="open_guide", help="Hướng dẫn sử dụng"):
            show_user_guide()
        render_account_menu(user, directory)


def render_audit_workspace(username: str = "") -> None:
    _init_session_state()
    with st.container(border=True, key="upload_box"):
        st.markdown(
            '<div class="section-title">📂 Báo cáo cần kiểm tra <span>file Word .docx</span></div>',
            unsafe_allow_html=True,
        )
        uploaded_file = st.file_uploader(
            "Kéo thả tệp DOCX vào đây hoặc bấm để chọn",
            type=["docx"],
            accept_multiple_files=False,
            help="Ứng dụng hiện xử lý một tệp DOCX cho mỗi lần kiểm tra.",
            label_visibility="collapsed",
        )
        hint_column, action_column = st.columns([3, 1], vertical_alignment="center")
        with hint_column:
            st.caption("Chọn tệp rồi bấm **Tạo báo cáo kiểm tra** để bắt đầu.")
        with action_column:
            process_requested = st.button(
                "Tạo báo cáo kiểm tra",
                type="primary",
                width="stretch",
                disabled=uploaded_file is None,
                help="Bấm để bắt đầu kiểm tra tệp DOCX đã chọn.",
            )

    if uploaded_file is None:
        st.info("Tải lên tệp BCTC dạng .docx để bắt đầu. Bấm ❓ ở góc phải trên cùng để xem hướng dẫn.")
        return

    uploaded_fingerprint = sha256(uploaded_file.getbuffer()).hexdigest()

    if process_requested:
        started_at = perf_counter()
        with st.spinner("Đang xử lý báo cáo..."):
            temp_path = None
            try:
                with NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
                    temp_file.write(uploaded_file.getbuffer())
                    temp_path = temp_file.name

                outputs = create_audit_outputs(temp_path)
            except UnsupportedFileError as exc:
                usage_log.log_event(
                    username, usage_log.EVENT_AUDIT_FAILED, f"{uploaded_file.name}; Định dạng không hỗ trợ"
                )
                st.error(str(exc))
                return
            except Exception:
                LOGGER.exception("Không xử lý được tệp báo cáo tải lên.")
                usage_log.log_event(username, usage_log.EVENT_AUDIT_FAILED, f"{uploaded_file.name}; Lỗi xử lý tệp")
                st.error("Không thể xử lý tệp. Vui lòng kiểm tra định dạng báo cáo và thử lại.")
                return
            finally:
                if temp_path:
                    Path(temp_path).unlink(missing_ok=True)

        st.session_state["tables"] = outputs.tables
        st.session_state["workbook_bytes"] = outputs.workbook_bytes
        st.session_state["commented_docx_bytes"] = outputs.commented_docx_bytes
        st.session_state["word_comment_failed"] = outputs.word_comment_failed
        st.session_state["attention_items"] = outputs.attention_items
        st.session_state["regime_detection"] = outputs.regime_detection
        st.session_state["applied_rule_pack"] = outputs.applied_rule_pack
        st.session_state["coverage"] = outputs.coverage
        st.session_state["processed_filename"] = uploaded_file.name
        st.session_state["processed_file_fingerprint"] = uploaded_fingerprint
        usage_log.log_event(
            username,
            usage_log.EVENT_AUDIT,
            _audit_log_details(uploaded_file.name, outputs, perf_counter() - started_at),
        )
        st.toast(f"Đã kiểm tra xong {len(outputs.tables)} bảng.", icon="✅")

    if st.session_state["word_comment_failed"]:
        st.warning(
            "Đã tạo file Excel, nhưng chưa thể tạo file Word có comment. "
            "Bạn vẫn có thể tải báo cáo Excel và thử lại với tệp DOCX khác."
        )

    if st.session_state["processed_file_fingerprint"] != uploaded_fingerprint:
        # Kết quả cũ thuộc tệp khác: không hiển thị để tránh nhầm báo cáo.
        st.info("Đã chọn tệp mới. Bấm **Tạo báo cáo kiểm tra** để xem kết quả.")
        return
    if st.session_state["workbook_bytes"]:
        _render_results()


def _apply_styles() -> None:
    st.markdown(f"<style>{STYLES_PATH.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


def _audit_log_details(filename: str, outputs: AuditOutputs, seconds: float) -> str:
    """Chi tiết nhật ký: tên tệp Word và số liệu vận hành; không chứa nội dung báo cáo."""
    differences = sum(1 for item in outputs.attention_items if item.severity == "Sai lệch")
    reviews = sum(1 for item in outputs.attention_items if item.severity == "Cần xem xét")
    return (
        f"{filename}; {len(outputs.tables)} bảng; {differences} sai lệch; {reviews} cần xem xét; "
        f"{seconds:.1f} giây"
    )


def _checked_output_stem(filename: str) -> str:
    source_stem = Path(filename or "bao_cao.docx").stem.strip()
    return f"{source_stem or 'bao_cao'} checked"


def _init_session_state() -> None:
    defaults = {
        "tables": [],
        "workbook_bytes": None,
        "commented_docx_bytes": None,
        "word_comment_failed": False,
        "attention_items": (),
        "regime_detection": None,
        "applied_rule_pack": None,
        "coverage": None,
        "processed_filename": "",
        "processed_file_fingerprint": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


SEVERITY_DIFFERENCE = "Sai lệch"
SEVERITY_REVIEW = "Cần xem xét"
WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _render_results() -> None:
    """Vùng kết quả gọn: một dòng kết luận, nút tải, danh sách sai lệch, nhóm cần xem xét thu gọn."""
    attention_items = st.session_state["attention_items"]
    regime_detection = st.session_state["regime_detection"]
    applied_rule_pack = st.session_state["applied_rule_pack"]
    coverage = st.session_state["coverage"]
    differences = [item for item in attention_items if item.severity == SEVERITY_DIFFERENCE]
    reviews = [item for item in attention_items if item.severity == SEVERITY_REVIEW]
    rules_applied = applied_rule_pack is not None and applied_rule_pack.version != "none"
    checked_stem = _checked_output_stem(st.session_state["processed_filename"])

    with st.container(border=True, key="result_box"):
        verdict_column, word_column, excel_column = st.columns(
            [3, 1.3, 0.9],
            vertical_alignment="center",
        )
        with verdict_column:
            st.markdown(
                _verdict_html(differences, reviews, rules_applied, coverage),
                unsafe_allow_html=True,
            )
        with word_column:
            if st.session_state["commented_docx_bytes"]:
                st.download_button(
                    "Tải Word có comment",
                    data=st.session_state["commented_docx_bytes"],
                    file_name=f"{checked_stem}.docx",
                    mime=WORD_MIME,
                    type="primary",
                    width="stretch",
                )
        with excel_column:
            if st.session_state["workbook_bytes"]:
                st.download_button(
                    "Excel chi tiết",
                    data=st.session_state["workbook_bytes"],
                    file_name=f"{checked_stem}.xlsx",
                    mime=EXCEL_MIME,
                    width="stretch",
                    help="Workbook có công thức kiểm tra cho từng bảng, dùng khi cần truy vết.",
                )
        st.markdown(_scope_line_html(coverage, applied_rule_pack), unsafe_allow_html=True)
        if not rules_applied:
            _render_rule_pack_status(regime_detection, applied_rule_pack)
        else:
            with st.expander("Chi tiết nhận diện Thông tư và phạm vi kiểm tra"):
                _render_rule_pack_status(regime_detection, applied_rule_pack)
                _render_coverage_caption(coverage)

        if not st.session_state["tables"]:
            st.info(
                "Không tìm thấy bảng nào trong tệp. Kiểm tra lại tệp có phải BCTC trình bày số liệu dạng bảng."
            )
        if differences:
            st.markdown(
                '<div class="finding-section danger">Sai lệch</div>'
                + "".join(_finding_row_html(item) for item in differences),
                unsafe_allow_html=True,
            )
        if reviews:
            with st.expander(
                f"Cần xem xét ({len(reviews)}) — {_review_preview(reviews)}",
                expanded=not differences,
            ):
                st.markdown(
                    "".join(_finding_row_html(item) for item in reviews),
                    unsafe_allow_html=True,
                )


def _verdict_html(
    differences: list[AuditAttentionItem],
    reviews: list[AuditAttentionItem],
    rules_applied: bool,
    coverage: AuditCoverage | None,
) -> str:
    chips: list[str] = []
    if differences:
        chips.append(f'<span class="verdict-chip danger">{len(differences)} sai lệch</span>')
    if reviews:
        chips.append(f'<span class="verdict-chip warning">{len(reviews)} cần xem xét</span>')
    if not chips:
        fully_covered = rules_applied and coverage is not None and not coverage.review_tables
        chips.append(
            '<span class="verdict-chip success">Không phát hiện sai lệch</span>'
            if fully_covered
            else '<span class="verdict-chip warning">Không phát hiện sai lệch trong phạm vi đã kiểm tra</span>'
        )
    return f'<div class="verdict-row">{"".join(chips)}</div>'


def _scope_line_html(coverage: AuditCoverage | None, applied_rule_pack) -> str:
    parts: list[str] = []
    if coverage is not None and coverage.total_tables:
        parts.append(f"Đã kiểm tra {coverage.checked_tables}/{coverage.total_tables} bảng")
        if coverage.review_tables:
            parts.append(f"{coverage.review_tables} bảng cần tự rà soát")
    if applied_rule_pack is not None and applied_rule_pack.version != "none":
        parts.append(escape(applied_rule_pack.display_name))
    return f'<div class="scope-line">{" · ".join(parts)}</div>' if parts else ""


def _finding_row_html(item: AuditAttentionItem) -> str:
    kind = "danger" if item.severity == SEVERITY_DIFFERENCE else "warning"
    detail = item.detail
    title = escape(item.location or item.table_title)
    tag_parts = [f"Bảng {item.table_index}", item.statement_label, item.currency]
    tag = escape(" · ".join(part for part in tag_parts if part))
    if detail is None:
        meta = escape(item.message)
        values = ""
        difference = "—"
    else:
        meta = escape(" · ".join(part for part in (detail.check, detail.period, detail.note) if part))
        values = escape(
            " · ".join(
                f"{label} {format_amount(value)}"
                for label, value in (
                    (detail.first_label, detail.first_value),
                    (detail.second_label, detail.second_value),
                )
                if value is not None
            )
        )
        if detail.difference is not None:
            difference = escape(format_amount(detail.difference))
        elif detail.first_value is not None:
            difference = "Sai dấu"
        else:
            difference = "—"
    values_html = f'<div class="finding-values">{values}</div>' if values else ""
    return (
        f'<div class="finding-row">'
        f'<div class="finding-main">'
        f'<div class="finding-title">{title}</div>'
        f'<div class="finding-meta"><span class="finding-tag">{tag}</span>{meta}</div>'
        f"{values_html}</div>"
        f'<div class="finding-diff {kind}">{difference}</div>'
        f"</div>"
    )


def _review_preview(reviews: list[AuditAttentionItem]) -> str:
    """Tóm tắt một dòng cho tiêu đề nhóm cần xem xét (tối đa hai mục)."""
    labels = [
        item.detail.check if item.detail and item.detail.check else (item.location or item.table_title)
        for item in reviews
    ]
    unique_labels = list(dict.fromkeys(labels))
    preview = " · ".join(unique_labels[:2])
    return preview + (" …" if len(unique_labels) > 2 else "")


def _render_coverage_caption(coverage: AuditCoverage | None) -> None:
    if coverage is None or not coverage.total_tables:
        return
    st.caption(
        f"Phạm vi kiểm tra tự động: {coverage.checked_tables}/{coverage.total_tables} bảng đã kiểm tra; "
        f"{coverage.review_tables} bảng chưa nhận diện được cấu trúc tổng, cần tự rà soát; "
        f"{coverage.not_applicable_tables} bảng thông tin không áp dụng kiểm tra số học. "
        "Chi tiết từng bảng xem tại sheet 00_Tong_hop của file Excel."
    )


def _render_rule_pack_status(regime_detection, applied_rule_pack) -> None:
    if regime_detection is None:
        st.warning("Chưa có thông tin nhận diện Thông tư và bộ rule áp dụng.")
        return

    if applied_rule_pack is None or applied_rule_pack.version == "none":
        st.warning(
            f"Thông tư nhận diện: {regime_detection.display_name}. "
            "Chưa có bộ rule nghiệp vụ phù hợp nên ứng dụng không áp dụng rule của Thông tư 200."
        )
    else:
        st.info(
            f"Bộ rule áp dụng: {applied_rule_pack.display_name} "
            f"(phiên bản {applied_rule_pack.version}). "
            f"Mức tin cậy nhận diện: {regime_detection.confidence.value}."
        )
    if regime_detection.message:
        st.caption(regime_detection.message)
    if regime_detection.evidence:
        st.caption("Căn cứ nhận diện: " + " | ".join(regime_detection.evidence))

from __future__ import annotations

from base64 import b64encode
from hashlib import sha256
from pathlib import Path
from tempfile import NamedTemporaryFile

import streamlit as st

from src.services.audit_service import (
    AuditOutputs,
    UnsupportedFileError,
    create_audit_outputs,
)


def main() -> None:
    st.set_page_config(
        page_title="Kiểm tra báo cáo tài chính",
        layout="wide",
    )

    _init_session_state()
    _apply_compact_styles()
    _render_sidebar_guide()

    header_col, link_col = st.columns([2.5, 1], vertical_alignment="bottom")
    with header_col:
        st.markdown("## Kiểm tra báo cáo tài chính")
        st.caption("Kéo thả hoặc chọn báo cáo DOCX để trích xuất bảng và tạo file Excel kiểm tra.")
    with link_col:
        st.markdown(
            """
            <div style="text-align: right; margin-bottom: 0.5rem;">
                <a href="https://wordnumcompareapp2-h4qbe8uc3ifcvpnwjldg7t.streamlit.app/" target="_blank" rel="noopener noreferrer" 
                   style="display: inline-block; padding: 0.45rem 0.85rem; background-color: #F1F5F9; color: #1F4E78; 
                          font-weight: 600; font-size: 0.85rem; border-radius: 0.45rem; border: 1px solid #CBD5E1; 
                          text-decoration: none; box-shadow: 0 1px 2px rgba(0,0,0,0.04); transition: all 0.2s ease;">
                   🌐 Đối chiếu số với bản dịch ↗
                </a>
            </div>
            """,
            unsafe_allow_html=True,
        )

    auto_process = st.checkbox(
        "Tự động tạo báo cáo kiểm tra sau khi có file Word",
        value=True,
        key="auto_process_after_upload",
        help="Tắt tùy chọn này nếu bạn muốn kiểm tra thiết lập trước rồi bấm nút tạo báo cáo.",
    )

    uploaded_file = st.file_uploader(
        "Kéo thả tệp DOCX vào đây hoặc bấm để chọn",
        type=["docx"],
        accept_multiple_files=False,
        help="Ứng dụng hiện xử lý một tệp DOCX cho mỗi lần kiểm tra.",
    )

    if uploaded_file is None:
        st.info("Vui lòng tải lên tệp DOCX để bắt đầu.")
        return

    _, action_column = st.columns([3, 1], vertical_alignment="bottom")
    with action_column:
        manual_process_requested = st.button(
            "Tạo báo cáo kiểm tra",
            type="primary",
            width="stretch",
            disabled=auto_process,
            help=(
                "Báo cáo sẽ tự động được tạo sau khi tải tệp."
                if auto_process
                else "Bấm để bắt đầu kiểm tra tệp DOCX đã chọn."
            ),
        )

    uploaded_fingerprint = sha256(uploaded_file.getbuffer()).hexdigest()
    process_requested = (
        auto_process
        and uploaded_fingerprint != st.session_state["processed_file_fingerprint"]
    ) or (not auto_process and manual_process_requested)

    if process_requested:
        with st.spinner("Đang xử lý báo cáo..."):
            temp_path = None
            try:
                with NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
                    temp_file.write(uploaded_file.getbuffer())
                    temp_path = temp_file.name

                outputs = create_audit_outputs(temp_path)
            except UnsupportedFileError as exc:
                st.error(str(exc))
                return
            except Exception:
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
        st.session_state["processed_filename"] = uploaded_file.name
        st.session_state["processed_file_fingerprint"] = uploaded_fingerprint
        st.success(f"Đã trích xuất {len(outputs.tables)} bảng và tạo báo cáo kiểm tra.")

    if st.session_state["word_comment_failed"]:
        st.warning(
            "Đã tạo file Excel, nhưng chưa thể tạo file Word có comment. "
            "Bạn vẫn có thể tải báo cáo Excel và thử lại với tệp DOCX khác."
        )

    checked_stem = _checked_output_stem(st.session_state["processed_filename"])
    combined_column, excel_column, word_column = st.columns(3)
    with combined_column:
        if st.session_state["workbook_bytes"] and st.session_state["commented_docx_bytes"]:
            _render_sequential_download_button(
                checked_stem,
                workbook_bytes=st.session_state["workbook_bytes"],
                commented_docx_bytes=st.session_state["commented_docx_bytes"],
            )

    with excel_column:
        if st.session_state["workbook_bytes"]:
            st.download_button(
                "Tải xuống Excel XLSX",
                data=st.session_state["workbook_bytes"],
                file_name=f"{checked_stem}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

    with word_column:
        if st.session_state["commented_docx_bytes"]:
            st.download_button(
                "Tải xuống Word có comment",
                data=st.session_state["commented_docx_bytes"],
                file_name=f"{checked_stem}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

    if st.session_state["tables"]:
        _render_attention_summary(
            st.session_state["attention_items"],
            st.session_state["regime_detection"],
            st.session_state["applied_rule_pack"],
        )


def _apply_compact_styles() -> None:
    st.markdown(
        """
        <style>
        .stMainBlockContainer {
            padding-top: 3.75rem;
            padding-bottom: 1.5rem;
        }
        [data-testid="stMarkdown"] h2 {
            font-size: 1.8rem;
            font-family: "Segoe UI", Arial, "Noto Sans", sans-serif;
            font-weight: 700;
            line-height: 1.45;
            letter-spacing: normal;
            overflow: visible;
            padding: 0.3rem 0 0.1rem;
            margin: 0 0 0.1rem;
        }
        [data-testid="stFileUploader"] {
            margin-top: -0.25rem;
        }
        [data-testid="stFileUploaderDropzone"] {
            min-height: 5.25rem;
            padding: 0.75rem 1rem;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] > div > span {
            font-size: 0.9rem;
        }
        [data-testid="stMetric"] {
            border: 1px solid rgba(128, 128, 128, 0.22);
            border-radius: 0.55rem;
            padding: 0.55rem 0.75rem;
        }
        [data-testid="stMetricLabel"] {
            min-height: auto;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.65rem;
            line-height: 1.25;
        }
        [data-testid="stAlert"] {
            padding: 0;
        }
        [data-testid="stAlertContainer"] {
            padding: 0.65rem 0.85rem;
        }
        [data-testid="stVerticalBlock"] {
            gap: 0.65rem;
        }
        header,
        [data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stToolbarActions"],
        [data-testid="stHeaderNav"],
        .stAppHeader,
        .stAppToolbar,
        #MainMenu,
        footer,
        a[href*="github.com"] {
            visibility: hidden !important;
            display: none !important;
            height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        [data-testid="stSidebar"] {
            background-color: #F8FAFC;
        }
        [data-testid="stSidebar"] [data-testid="stMarkdown"] h3 {
            font-size: 1.15rem;
            color: #1F4E78;
            font-weight: 700;
            margin-bottom: 0.15rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_sidebar_guide() -> None:
    with st.sidebar:
        st.markdown("### 📖 Hướng dẫn 10 bước")
        st.caption("Quy trình kiểm tra BCTC cho kiểm toán viên")

        guide_steps = [
            ("1. Chuẩn bị tệp nguồn", "Sử dụng tệp Báo cáo tài chính định dạng Word `.docx` rõ ràng."),
            ("2. Tải tệp lên ứng dụng", "Kéo thả tệp `.docx` vào vùng tải lên ở giao diện chính hoặc nhấn chọn file."),
            ("3. Tự động trích xuất", "Hệ thống tự động đọc bảng và lập báo cáo ngay khi tải tệp lên."),
            ("4. Nhận diện Thông tư", "Tự động nhận diện chế độ kế toán (TT 200 / TT 133) và áp dụng RulePack."),
            ("5. Xem tổng quan kết quả", "Theo dõi tổng số **🔴 Sai lệch**, **🟠 Cần xem xét** và danh sách bảng bị ảnh hưởng."),
            ("6. Tải xuống kết quả", "Nhấn **Tải xuống Word và Excel** (hoặc tải từng file XLSX / DOCX)."),
            ("7. Thứ tự Sheet Excel", "Các sheet có số liệu (`00_Tong_hop`, BS, PL, CF, TM) xếp trước; chữ ký/thông tin xếp sau."),
            ("8. Màu sắc Tab Sheet", "Tab **🔴 Đỏ sẫm (`#C00000`)**: Sai lệch; **🟡 Vàng (`#FFC000`)**: Cần xem xét; **🟢 Xanh (`#70AD47`)**: Khớp."),
            ("9. Điều hướng nhanh Excel", "Dùng ô điều hướng A1/A2/B2 và phím tắt `Ctrl + Page Up/Down` để chuyển sheet."),
            ("10. Rà soát Word Comment", "Mở file Word kết quả để xem ghi chú comment tự động đánh dấu vị trí chênh lệch."),
        ]

        with st.expander("📋 Danh sách 10 bước chi tiết", expanded=True):
            for step_title, step_desc in guide_steps:
                st.markdown(
                    f"<div style='margin-bottom: 0.45rem; padding: 0.45rem 0.6rem; background: #FFFFFF; "
                    f"border: 1px solid #E2E8F0; border-radius: 0.4rem; box-shadow: 0 1px 2px rgba(0,0,0,0.02);'>"
                    f"<strong style='color: #1F4E78; font-size: 0.85rem;'>{step_title}</strong><br/>"
                    f"<span style='color: #475569; font-size: 0.78rem; line-height: 1.3;'>{step_desc}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

        st.markdown(
            """
            <div style="margin-top: 0.5rem;">
                <a href="https://wordnumcompareapp2-h4qbe8uc3ifcvpnwjldg7t.streamlit.app/" target="_blank" rel="noopener noreferrer" 
                   style="display: block; text-align: center; padding: 0.5rem; background: #EFF6FF; color: #1D4ED8; 
                          font-weight: 600; font-size: 0.82rem; border-radius: 0.45rem; border: 1px solid #BFDBFE; text-decoration: none;
                          box-shadow: 0 1px 2px rgba(0,0,0,0.03);">
                   🌐 Đối chiếu số với bản dịch ↗
                </a>
            </div>
            """,
            unsafe_allow_html=True,
        )


def _checked_output_stem(filename: str) -> str:
    source_stem = Path(filename or "bao_cao.docx").stem.strip()
    return f"{source_stem or 'bao_cao'} checked"


def _render_sequential_download_button(
    checked_stem: str,
    workbook_bytes: bytes,
    commented_docx_bytes: bytes,
) -> None:
    word_data = b64encode(commented_docx_bytes).decode("ascii")
    excel_data = b64encode(workbook_bytes).decode("ascii")
    word_name = b64encode(f"{checked_stem}.docx".encode("utf-8")).decode("ascii")
    excel_name = b64encode(f"{checked_stem}.xlsx".encode("utf-8")).decode("ascii")
    st.components.v1.html(
        f"""
        <button id="download-both" type="button">Tải xuống Word và Excel</button>
        <script>
        const button = document.getElementById("download-both");
        const decodeName = value => new TextDecoder().decode(
            Uint8Array.from(atob(value), character => character.charCodeAt(0))
        );
        const download = (data, mime, encodedName) => {{
            const bytes = Uint8Array.from(atob(data), character => character.charCodeAt(0));
            const url = URL.createObjectURL(new Blob([bytes], {{type: mime}}));
            const link = document.createElement("a");
            link.href = url;
            link.download = decodeName(encodedName);
            document.body.appendChild(link);
            link.click();
            link.remove();
            window.setTimeout(() => URL.revokeObjectURL(url), 10000);
        }};
        button.addEventListener("click", async () => {{
            button.disabled = true;
            download("{word_data}",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "{word_name}");
            await new Promise(resolve => window.setTimeout(resolve, 700));
            download("{excel_data}",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "{excel_name}");
            button.disabled = false;
        }});
        </script>
        <style>
        html, body {{ margin: 0; padding: 0; background: transparent; }}
        #download-both {{
            width: 100%; height: 40px; padding: 0 0.75rem;
            color: white; background: linear-gradient(135deg, #1F4E78 0%, #2563EB 100%);
            border: none; border-radius: 0.5rem;
            font: 600 0.95rem/1.2 "Segoe UI", Arial, sans-serif; cursor: pointer;
            box-shadow: 0 2px 6px rgba(31, 78, 120, 0.22);
            transition: all 0.2s ease;
        }}
        #download-both:hover {{ opacity: 0.92; box-shadow: 0 4px 10px rgba(31, 78, 120, 0.35); }}
        #download-both:disabled {{ cursor: wait; opacity: 0.65; }}
        </style>
        """,
        height=41,
        scrolling=False,
    )


def _init_session_state() -> None:
    defaults = {
        "tables": [],
        "workbook_bytes": None,
        "commented_docx_bytes": None,
        "word_comment_failed": False,
        "attention_items": (),
        "regime_detection": None,
        "applied_rule_pack": None,
        "processed_filename": "",
        "processed_file_fingerprint": "",
        "auto_process_after_upload": True,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _render_attention_summary(attention_items: tuple, regime_detection=None, applied_rule_pack=None) -> None:
    st.subheader("Các vấn đề cần chú ý")
    _render_rule_pack_status(regime_detection, applied_rule_pack)

    differences = [item for item in attention_items if item.severity == "Sai lệch"]
    reviews = [item for item in attention_items if item.severity == "Cần xem xét"]
    affected_tables = len({item.table_index for item in attention_items})

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🔴 Sai lệch", len(differences))
    with col2:
        st.metric("🟠 Cần xem xét", len(reviews))
    with col3:
        st.metric("📋 Bảng bị ảnh hưởng", affected_tables)

    if attention_items:
        tab_diff, tab_rev = st.tabs([
            f"🔴 Sai lệch ({len(differences)})",
            f"🟠 Cần xem xét ({len(reviews)})",
        ])

        def _render_item_table(items_list: list) -> None:
            if not items_list:
                st.info("Không có ghi nhận trong danh mục này.")
                return
            issue_rows = [
                {
                    "Mức độ": "🔴 Sai lệch" if item.severity == "Sai lệch" else "🟠 Cần xem xét",
                    "Bảng": item.table_index,
                    "Nội dung bảng": item.table_title,
                    "Loại bảng": item.table_type,
                    "Vấn đề": item.message,
                }
                for item in items_list
            ]
            table_height = min(400, 38 * (len(issue_rows) + 1) + 4)
            st.dataframe(issue_rows, width="stretch", height=table_height, hide_index=True)

        with tab_diff:
            _render_item_table(differences)
        with tab_rev:
            _render_item_table(reviews)
    else:
        st.success("Không phát hiện sai lệch hoặc trường hợp cần xem xét trong các bảng đã xử lý.")


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

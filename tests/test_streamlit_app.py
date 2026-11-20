from decimal import Decimal
from contextlib import nullcontext
import importlib
import sys
from types import ModuleType
from types import SimpleNamespace
import unittest
from unittest.mock import patch


class _UploadedFile:
    name = "bao_cao.docx"

    def getbuffer(self):
        return b"docx"


class _FakeSidebar:
    def __init__(self, parent):
        self.parent = parent
        self.markdown_calls = []

    def __enter__(self):
        self.parent._in_sidebar = True
        return self

    def __exit__(self, *args):
        self.parent._in_sidebar = False

    def title(self, *args, **kwargs):
        pass

    def header(self, *args, **kwargs):
        pass

    def subheader(self, *args, **kwargs):
        pass

    def markdown(self, *args, **kwargs):
        self.markdown_calls.append((args, kwargs))

    def caption(self, *args, **kwargs):
        pass

    def expander(self, label, **kwargs):
        return nullcontext()


class _FakeStreamlit(ModuleType):
    def __init__(self):
        super().__init__("streamlit")
        self.session_state = {}
        self._in_sidebar = False
        self.sidebar = _FakeSidebar(self)
        self.selectbox_calls = []
        self.checkbox_calls = []
        self.download_calls = []
        self.metric_calls = []
        self.dataframe_calls = []
        self.success_calls = []
        self.info_calls = []
        self.warning_calls = []
        self.file_uploader_calls = []
        self.markdown_calls = []
        self.button_calls = []
        self.component_html_calls = []
        self.toast_calls = []
        self.expander_labels = []
        self.components = SimpleNamespace(
            v1=SimpleNamespace(html=self._component_html)
        )

    @staticmethod
    def _cache_decorator(func=None, **kwargs):
        def wrap(inner):
            inner.clear = lambda: None
            return inner

        return wrap(func) if callable(func) else wrap

    cache_data = _cache_decorator
    cache_resource = _cache_decorator

    def _component_html(self, *args, **kwargs):
        self.component_html_calls.append((args, kwargs))

    def set_page_config(self, **kwargs):
        pass

    def title(self, *args, **kwargs):
        pass

    def write(self, *args, **kwargs):
        pass

    def markdown(self, *args, **kwargs):
        if self._in_sidebar:
            self.sidebar.markdown_calls.append((args, kwargs))
        else:
            self.markdown_calls.append((args, kwargs))

    def file_uploader(self, *args, **kwargs):
        self.file_uploader_calls.append((args, kwargs))
        return _UploadedFile()

    def subheader(self, *args, **kwargs):
        pass

    def columns(self, spec, **kwargs):
        count = spec if isinstance(spec, int) else len(spec)
        return [nullcontext() for _ in range(count)]

    def selectbox(self, label, options, *, format_func, key, **kwargs):
        selected = self.session_state.get(key, options[0])
        self.session_state[key] = selected
        self.selectbox_calls.append((label, list(options), format_func(selected), key))
        return selected

    def caption(self, *args, **kwargs):
        pass

    def checkbox(self, label, *, value, key, **kwargs):
        selected = self.session_state.get(key, value)
        self.session_state[key] = selected
        self.checkbox_calls.append((label, value, key))
        return selected

    def button(self, *args, **kwargs):
        self.button_calls.append((args, kwargs))
        return True

    def spinner(self, *args, **kwargs):
        return nullcontext()

    def error(self, *args, **kwargs):
        pass

    def success(self, *args, **kwargs):
        self.success_calls.append((args, kwargs))

    def download_button(self, *args, **kwargs):
        self.download_calls.append((args, kwargs))

    def dataframe(self, *args, **kwargs):
        self.dataframe_calls.append((args, kwargs))

    def metric(self, *args, **kwargs):
        self.metric_calls.append((args, kwargs))

    def info(self, *args, **kwargs):
        self.info_calls.append((args, kwargs))

    def warning(self, *args, **kwargs):
        self.warning_calls.append((args, kwargs))

    def toast(self, *args, **kwargs):
        self.toast_calls.append((args, kwargs))

    def tabs(self, spec):
        return [nullcontext() for _ in range(len(spec))]

    def expander(self, label, **kwargs):
        self.expander_labels.append((label, kwargs))
        return nullcontext()

    def container(self, *args, **kwargs):
        return nullcontext()

    def rerun(self):
        pass


class StreamlitInterfaceTest(unittest.TestCase):
    def tearDown(self):
        sys.modules.pop("src.ui.streamlit_app", None)

    def test_processing_uses_default_workbook_order_without_sort_controls(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            with patch.object(
                streamlit_app,
                "create_audit_outputs",
                return_value=streamlit_app.AuditOutputs(
                    tables=[],
                    workbook_bytes=b"workbook",
                    commented_docx_bytes=b"word",
                ),
            ) as create_audit_outputs:
                streamlit_app.render_audit_workspace("VPSG")

        self.assertEqual(fake_streamlit.selectbox_calls, [])
        self.assertIn("Kéo thả", fake_streamlit.file_uploader_calls[0][0][0])
        section_title = fake_streamlit.markdown_calls[0][0][0]
        self.assertIn("Báo cáo cần kiểm tra", section_title)
        # Không còn tùy chọn tự động xử lý; người dùng bấm nút để tạo báo cáo.
        self.assertEqual(fake_streamlit.checkbox_calls, [])
        self.assertEqual(fake_streamlit.button_calls[0][0][0], "Tạo báo cáo kiểm tra")
        self.assertFalse(fake_streamlit.button_calls[0][1]["disabled"])
        create_audit_outputs.assert_called_once()
        self.assertEqual(create_audit_outputs.call_args.kwargs, {})
        self.assertEqual(
            [(call[0][0], call[1]["file_name"], call[1].get("type")) for call in fake_streamlit.download_calls],
            [
                ("Tải Word có comment", "bao_cao checked.docx", "primary"),
                ("Excel chi tiết", "bao_cao checked.xlsx", None),
            ],
        )
        # Không còn nút tải đồng thời Word và Excel qua component HTML.
        self.assertEqual(fake_streamlit.component_html_calls, [])

    def test_upload_alone_does_not_start_processing(self):
        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            with patch.object(streamlit_app, "create_audit_outputs") as create_audit_outputs:
                streamlit_app.render_audit_workspace()

        create_audit_outputs.assert_not_called()
        self.assertIn("Tạo báo cáo kiểm tra", fake_streamlit.info_calls[-1][0][0])
        self.assertEqual(fake_streamlit.download_calls, [])

    def test_button_processes_selected_file(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            with patch.object(
                streamlit_app,
                "create_audit_outputs",
                return_value=streamlit_app.AuditOutputs(tables=[], workbook_bytes=b"workbook"),
            ) as create_audit_outputs:
                streamlit_app.render_audit_workspace()

        create_audit_outputs.assert_called_once()

    def test_results_of_previous_file_are_hidden_for_new_upload(self):
        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False
        fake_streamlit.session_state.update(
            {
                "tables": [object()],
                "workbook_bytes": b"old",
                "commented_docx_bytes": b"old",
                "processed_file_fingerprint": "fingerprint-of-another-file",
            }
        )
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            streamlit_app.render_audit_workspace()

        self.assertEqual(fake_streamlit.download_calls, [])
        self.assertIn("Đã chọn tệp mới", fake_streamlit.info_calls[-1][0][0])

    def test_checked_output_stem_preserves_uploaded_word_name(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")

        self.assertEqual(
            streamlit_app._checked_output_stem("Báo cáo tài chính 2025.docx"),
            "Báo cáo tài chính 2025 checked",
        )

    @staticmethod
    def _render_results(fake_streamlit, items, *, coverage=None):
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            from src.domain.audit_rules import TT200_RULE_PACK

            streamlit_app._init_session_state()
            fake_streamlit.session_state.update(
                {
                    "tables": [object()],
                    "attention_items": tuple(items),
                    "applied_rule_pack": TT200_RULE_PACK,
                    "coverage": coverage,
                    "workbook_bytes": b"x",
                    "commented_docx_bytes": b"w",
                    "processed_filename": "bao_cao.docx",
                }
            )
            streamlit_app._render_results()
        return " ".join(str(call[0][0]) for call in fake_streamlit.markdown_calls)

    def test_results_show_differences_with_structured_amounts(self):
        from src.domain.models import FindingDetail
        from src.services.audit_service import AuditAttentionItem

        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False
        item = AuditAttentionItem(
            "Sai lệch", 54, "Báo cáo lưu chuyển tiền tệ", "Báo cáo lưu chuyển tiền tệ", "m",
            "Lưu chuyển tiền thuần trong kỳ",
            FindingDetail(
                check="Mã 50 = 20 + 30 + 40",
                period="kỳ này",
                first_label="Trình bày",
                first_value=Decimal("48408766679"),
                second_label="Tính lại",
                second_value=Decimal("48408766678"),
                difference=Decimal("1"),
            ),
            "CF",
            "VND",
        )

        rendered = self._render_results(fake_streamlit, [item])

        self.assertIn('<span class="verdict-chip danger">1 sai lệch</span>', rendered)
        self.assertIn("Lưu chuyển tiền thuần trong kỳ", rendered)
        self.assertIn("Bảng 54 · CF · VND", rendered)
        self.assertIn("Mã 50 = 20 + 30 + 40 · kỳ này", rendered)
        self.assertIn("Trình bày 48.408.766.679 · Tính lại 48.408.766.678", rendered)
        self.assertIn('<div class="finding-diff danger">1</div>', rendered)
        self.assertEqual(fake_streamlit.metric_calls, [])
        self.assertEqual(fake_streamlit.dataframe_calls, [])

    def test_reviews_are_collapsed_when_differences_exist(self):
        from src.domain.models import FindingDetail
        from src.services.audit_service import AuditAttentionItem

        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False
        items = [
            AuditAttentionItem("Sai lệch", 1, "Thuế", "Thuyết minh", "m", "Tổng cộng"),
            AuditAttentionItem(
                "Cần xem xét", 2, "BS", "BS", "m", "Tài sản ngắn hạn khác",
                FindingDetail(check="Mã 150 = 151 + … + 155 (5 mã)", note="Mã 153 không phải số"),
            ),
        ]

        self._render_results(fake_streamlit, items)

        review_expanders = [
            (label, kwargs)
            for label, kwargs in fake_streamlit.expander_labels
            if label.startswith("Cần xem xét")
        ]
        self.assertEqual(len(review_expanders), 1)
        self.assertIn("Mã 150 = 151 + … + 155 (5 mã)", review_expanders[0][0])
        self.assertFalse(review_expanders[0][1]["expanded"])

    def test_clean_result_with_unchecked_tables_is_not_reported_as_fully_matched(self):
        from src.services.audit_service import AuditCoverage

        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False

        rendered = self._render_results(
            fake_streamlit,
            [],
            coverage=AuditCoverage(total_tables=5, checked_tables=3, review_tables=2),
        )

        self.assertIn("Không phát hiện sai lệch trong phạm vi đã kiểm tra", rendered)
        self.assertIn("Đã kiểm tra 3/5 bảng · 2 bảng cần tự rà soát", rendered)
        self.assertNotIn("verdict-chip success", rendered)

    def test_results_escape_document_text(self):
        from src.services.audit_service import AuditAttentionItem

        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False

        rendered = self._render_results(
            fake_streamlit,
            [AuditAttentionItem("Sai lệch", 1, "<script>x</script>", "Thuyết minh", "a < b")],
        )

        self.assertNotIn("<script>", rendered)
        self.assertIn("&lt;script&gt;", rendered)
        self.assertIn("a &lt; b", rendered)

    def test_user_guide_covers_workflow_results_and_account(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            user_guide = importlib.import_module("src.ui.user_guide")
            sections = dict(user_guide.guide_sections())
        sys.modules.pop("src.ui.user_guide", None)

        self.assertIn("🚀 Bắt đầu nhanh", sections)
        self.assertIn("Thông tư 200", sections["📂 Kiểm tra báo cáo"])
        self.assertIn("00_Tong_hop", sections["📊 Đọc kết quả Excel"])
        self.assertIn("CHECK_FS - Sai lệch", sections["💬 Rà soát Word"])
        self.assertIn("Đổi mật khẩu", sections["👤 Tài khoản"])

    def test_styles_file_contains_shared_ui_classes(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            streamlit_app._apply_styles()

        css = fake_streamlit.markdown_calls[0][0][0]
        for selector in (
            "verdict-chip", "finding-row", "finding-diff", "st-key-top_bar", "st-key-login_box",
            "st-key-upload_box", "auth-wait", "stSidebar", "Inter",
        ):
            self.assertIn(selector, css)

    def test_audit_log_details_include_word_file_name(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            from src.services.audit_service import AuditAttentionItem

            details = streamlit_app._audit_log_details(
                "BCTC 2025.docx",
                streamlit_app.AuditOutputs(
                    tables=[object(), object()],
                    workbook_bytes=b"x",
                    attention_items=(AuditAttentionItem("Sai lệch", 1, "Thuế", "Thuyết minh", "m"),),
                ),
                1.25,
            )

        self.assertEqual(details, "BCTC 2025.docx; 2 bảng; 1 sai lệch; 0 cần xem xét; 1.2 giây")


if __name__ == "__main__":
    unittest.main()

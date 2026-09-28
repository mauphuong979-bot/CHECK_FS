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
        self.components = SimpleNamespace(
            v1=SimpleNamespace(html=self._component_html)
        )

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

    def tabs(self, spec):
        return [nullcontext() for _ in range(len(spec))]

    def expander(self, label, **kwargs):
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
                streamlit_app.main()

        self.assertEqual(fake_streamlit.selectbox_calls, [])
        self.assertIn("Kéo thả", fake_streamlit.file_uploader_calls[0][0][0])
        compact_css = fake_streamlit.markdown_calls[0][0][0]
        rendered_title = fake_streamlit.markdown_calls[1][0][0]
        self.assertIn('padding-top: 3.75rem', compact_css)
        self.assertIn('font-family: "Segoe UI", Arial', compact_css)
        self.assertIn('[data-testid="stAlertContainer"]', compact_css)
        self.assertEqual(rendered_title, "## Kiểm tra báo cáo tài chính")
        self.assertEqual(
            [call[:3] for call in fake_streamlit.checkbox_calls],
            [
                ("Tự động tạo báo cáo kiểm tra sau khi có file Word", True, "auto_process_after_upload"),
            ],
        )
        self.assertTrue(fake_streamlit.button_calls[0][1]["disabled"])
        create_audit_outputs.assert_called_once()
        self.assertEqual(create_audit_outputs.call_args.kwargs, {})
        self.assertEqual(len(fake_streamlit.download_calls), 2)
        self.assertEqual(
            [call[1]["file_name"] for call in fake_streamlit.download_calls],
            ["bao_cao checked.xlsx", "bao_cao checked.docx"],
        )
        self.assertEqual(len(fake_streamlit.component_html_calls), 1)
        combined_html = fake_streamlit.component_html_calls[0][0][0]
        self.assertIn("Tải xuống Word và Excel", combined_html)
        self.assertLess(combined_html.index("wordprocessingml"), combined_html.index("spreadsheetml"))
        self.assertNotIn(".zip", combined_html)

    def test_auto_processing_only_runs_once_for_the_same_uploaded_file(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            with patch.object(
                streamlit_app,
                "create_audit_outputs",
                return_value=streamlit_app.AuditOutputs(tables=[], workbook_bytes=b"workbook"),
            ) as create_audit_outputs:
                streamlit_app.main()
                streamlit_app.main()

        create_audit_outputs.assert_called_once()

    def test_manual_button_processes_when_auto_processing_is_disabled(self):
        fake_streamlit = _FakeStreamlit()
        fake_streamlit.session_state["auto_process_after_upload"] = False
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            with patch.object(
                streamlit_app,
                "create_audit_outputs",
                return_value=streamlit_app.AuditOutputs(tables=[], workbook_bytes=b"workbook"),
            ) as create_audit_outputs:
                streamlit_app.main()

        self.assertFalse(fake_streamlit.button_calls[0][1]["disabled"])
        create_audit_outputs.assert_called_once()

    def test_checked_output_stem_preserves_uploaded_word_name(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")

        self.assertEqual(
            streamlit_app._checked_output_stem("Báo cáo tài chính 2025.docx"),
            "Báo cáo tài chính 2025 checked",
        )

    def test_attention_summary_only_renders_actionable_issues(self):
        fake_streamlit = _FakeStreamlit()
        fake_streamlit.button = lambda *args, **kwargs: False
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            streamlit_app._init_session_state()
            from src.services.audit_service import AuditAttentionItem

            streamlit_app._render_attention_summary(
                (
                    AuditAttentionItem(
                        "Sai lệch", 1, "Thuế", "Thuyết minh", "Có chênh lệch 10 VND."
                    ),
                ),
            )

        self.assertEqual([call[0] for call in fake_streamlit.metric_calls], [("🔴 Sai lệch", 1), ("🟠 Cần xem xét", 0), ("📋 Bảng bị ảnh hưởng", 1)])
        self.assertEqual(len(fake_streamlit.dataframe_calls), 1)
        self.assertEqual(fake_streamlit.dataframe_calls[0][0][0][0]["Mức độ"], "🔴 Sai lệch")

    def test_sidebar_renders_10_step_guide(self):
        fake_streamlit = _FakeStreamlit()
        with patch.dict(sys.modules, {"streamlit": fake_streamlit}):
            streamlit_app = importlib.import_module("src.ui.streamlit_app")
            streamlit_app._render_sidebar_guide()

        sidebar_markdowns = [call[0][0] for call in fake_streamlit.sidebar.markdown_calls]
        self.assertTrue(any("### 📖 Hướng dẫn 10 bước" in md for md in sidebar_markdowns))
        self.assertTrue(any("1. Chuẩn bị tệp nguồn" in md for md in sidebar_markdowns))
        self.assertTrue(any("10. Rà soát Word Comment" in md for md in sidebar_markdowns))


if __name__ == "__main__":
    unittest.main()


if __name__ == "__main__":
    unittest.main()

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from docx import Document

from src.domain.models import WordCommentCategory, WordCommentTarget
from src.reporting.word_comment_report import build_commented_docx


class WordCommentReportTest(unittest.TestCase):
    def test_adds_true_word_comment_and_aggregates_same_cell(self):
        with TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.docx"
            document = Document()
            table = document.add_table(rows=1, cols=2)
            table.cell(0, 0).text = "Chỉ tiêu"
            table.cell(0, 1).text = "100"
            document.save(source_path)

            output = build_commented_docx(
                source_path,
                [
                    WordCommentTarget(1, 0, 1, "Có chênh lệch."),
                    WordCommentTarget(1, 0, 1, "Cần kiểm tra."),
                ],
            )

        reopened = Document(BytesIO(output))
        comments = list(reopened.comments)
        self.assertEqual(len(comments), 1)
        self.assertIn("Có chênh lệch.", comments[0].text)
        self.assertIn("Cần kiểm tra.", comments[0].text)
        with ZipFile(BytesIO(output)) as archive:
            names = set(archive.namelist())
            document_xml = archive.read("word/document.xml")
        self.assertIn("word/comments.xml", names)
        self.assertIn(b"commentRangeStart", document_xml)
        self.assertIn(b"commentRangeEnd", document_xml)
        self.assertIn(b"commentReference", document_xml)

    def test_empty_target_falls_back_without_adding_visible_text(self):
        with TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.docx"
            document = Document()
            table = document.add_table(rows=1, cols=2)
            table.cell(0, 0).text = "Nhãn"
            table.cell(0, 1).text = ""
            document.save(source_path)

            output = build_commented_docx(
                source_path,
                [WordCommentTarget(1, 0, 1, "Không tìm thấy.")],
            )

        reopened = Document(BytesIO(output))
        self.assertEqual(reopened.tables[0].cell(0, 1).text, "")
        comments = list(reopened.comments)
        self.assertEqual(len(comments), 1)
        self.assertEqual(comments[0].text, "Không tìm thấy.")
        self.assertEqual(comments[0].author, "CHECK_FS")

    def test_separates_difference_and_review_comments_by_author_and_prefix(self):
        with TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.docx"
            document = Document()
            table = document.add_table(rows=1, cols=1)
            table.cell(0, 0).text = "100"
            document.save(source_path)

            output = build_commented_docx(
                source_path,
                [
                    WordCommentTarget(
                        1, 0, 0, "Cần kiểm toán viên xác minh.",
                        WordCommentCategory.REVIEW,
                    ),
                    WordCommentTarget(
                        1, 0, 0, "Đối chiếu TM–BS: chênh lệch 10.",
                        WordCommentCategory.DIFFERENCE,
                    ),
                ],
            )

        comments = list(Document(BytesIO(output)).comments)
        self.assertEqual(len(comments), 2)
        difference, review = comments
        self.assertEqual(difference.author, "CHECK_FS - Sai lệch")
        self.assertEqual(difference.text, "[SAI LỆCH] Đối chiếu TM–BS: chênh lệch 10.")
        difference_runs = [run for paragraph in difference.paragraphs for run in paragraph.runs]
        self.assertEqual(str(difference_runs[0].font.color.rgb), "C00000")
        self.assertTrue(difference_runs[0].bold)
        self.assertEqual(str(difference_runs[1].font.color.rgb), "0563C1")
        self.assertEqual(str(difference_runs[-1].font.color.rgb), "C00000")
        self.assertEqual(review.author, "CHECK_FS - Cần xem xét")
        self.assertEqual(review.text, "[CẦN XEM XÉT] Cần kiểm toán viên xác minh.")
        review_runs = [run for paragraph in review.paragraphs for run in paragraph.runs]
        self.assertEqual(str(review_runs[-1].font.color.rgb), "C65911")

    def test_empty_row_does_not_borrow_cell_from_other_rows(self):
        with TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.docx"
            document = Document()
            table = document.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "Tiêu đề"
            table.cell(0, 1).text = "Số cuối năm"
            document.save(source_path)

            output = build_commented_docx(
                source_path,
                [WordCommentTarget(1, 1, 1, "Không neo.", WordCommentCategory.REVIEW)],
            )

        self.assertEqual(list(Document(BytesIO(output)).comments), [])

if __name__ == "__main__":
    unittest.main()

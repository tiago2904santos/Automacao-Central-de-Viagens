from docx import Document
from docx.shared import Pt
from django.test import SimpleTestCase

from viagens.documents.document import replace_placeholders_in_paragraph


class DocxReplaceTests(SimpleTestCase):
    def test_replace_preserves_run_formatting(self):
        doc = Document()
        paragraph = doc.add_paragraph()
        run = paragraph.add_run("RG: {{rg}}")
        run.bold = True
        run.font.size = Pt(10.5)

        replace_placeholders_in_paragraph(paragraph, {"rg": "123"})

        self.assertEqual(paragraph.runs[0].text, "RG: 123")
        self.assertTrue(paragraph.runs[0].bold)
        self.assertEqual(paragraph.runs[0].font.size, Pt(10.5))

    def test_replace_across_runs_keeps_styles(self):
        doc = Document()
        paragraph = doc.add_paragraph()
        first = paragraph.add_run("CPF: {{")
        first.bold = True
        second = paragraph.add_run("cpf}}")
        second.bold = True

        replace_placeholders_in_paragraph(paragraph, {"cpf": "999"})

        self.assertEqual(paragraph.runs[0].text, "CPF: 999")
        self.assertTrue(paragraph.runs[0].bold)
        self.assertTrue(paragraph.runs[1].bold)

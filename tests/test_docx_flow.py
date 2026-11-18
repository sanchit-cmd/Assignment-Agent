import tempfile
import unittest
from pathlib import Path

import docx

from config import AgentConfig
from docx_parser import DocxParser
from docx_writer import DocxWriter
from gemini_solver import GeminiSolver, SolvedWorksheet


class DocxFlowTests(unittest.TestCase):
    def create_worksheet(self, directory: str, include_answers: bool = False) -> Path:
        document = docx.Document()
        document.add_paragraph("Topic: Example Topic\tWorksheet No. 1")
        for index in range(11):
            document.add_paragraph(f"Front matter {index}")
        document.add_paragraph("Example Topic")
        document.add_paragraph("Name:")
        document.add_paragraph("Roll No.:")
        document.add_paragraph("Questions")
        document.add_paragraph("Q7. Put ______ here.")
        document.add_paragraph("Q9. Explain the result.")
        if include_answers:
            document.add_paragraph("Answers:")

        path = Path(directory) / "worksheet.docx"
        document.save(path)
        return path

    def test_parser_detects_late_metadata_and_skips_title(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parsed = DocxParser.parse(self.create_worksheet(directory))

        self.assertEqual(parsed.student_info_mode, "multi")
        self.assertIn("reg_no", parsed.student_metadata_map)
        self.assertEqual([question.label for question in parsed.questions], ["Q7", "Q9"])
        self.assertTrue(all("Name:" not in question.text for question in parsed.questions))
        self.assertTrue(all("Example Topic" not in question.text for question in parsed.questions))

    def test_parser_keeps_combined_metadata_detection(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            document = docx.Document()
            document.add_paragraph("Topic: Example Topic\tWorksheet No. 1")
            for index in range(12):
                document.add_paragraph(f"Front matter {index}")
            document.add_paragraph("Name: \t Reg.No.: \t Date:")
            document.add_paragraph("Questions")
            document.add_paragraph("Q4. Explain this concept.")
            path = Path(directory) / "combined.docx"
            document.save(path)

            parsed = DocxParser.parse(path)
            output = DocxWriter(
                AgentConfig(date_str="01/01/2026", output_dir=Path(directory))
            ).update_document(
                parsed,
                SolvedWorksheet(),
                Path(directory) / "combined-output.docx",
            )
            result = docx.Document(output)

        self.assertEqual(parsed.student_info_mode, "single")
        self.assertIn("Q4", [question.label for question in parsed.questions])
        metadata_text = next(
            paragraph.text for paragraph in result.paragraphs if paragraph.text.startswith("Name:")
        )
        self.assertIn("Name: Sanchit Nigam", metadata_text)
        self.assertIn("Reg.No.: RA2311003030178", metadata_text)

    def test_writer_fills_roll_number_and_formats_blank_answer(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self.create_worksheet(directory)
            parsed = DocxParser.parse(source)
            blank_paragraph = parsed.questions[0].paragraph_indices[0]
            solved = SolvedWorksheet(
                blank_answers={blank_paragraph: ["filled"]},
                qa_answers=[
                    {
                        "q_index": 1,
                        "label": "Q9",
                        "question": parsed.questions[1].text,
                        "answer": "Explained.",
                    }
                ],
            )
            config = AgentConfig(date_str="01/01/2026", output_dir=Path(directory))
            output = DocxWriter(config).update_document(
                parsed, solved, Path(directory) / "solved.docx"
            )
            result = docx.Document(output)

        text = [paragraph.text for paragraph in result.paragraphs]
        self.assertIn("Roll No.: RA2311003030178", text)
        self.assertIn("ANSWERS", text)
        self.assertIn("9. Explained.", text)
        answer_run = next(
            run
            for paragraph in result.paragraphs
            for run in paragraph.runs
            if run.text == "filled"
        )
        self.assertTrue(answer_run.bold)
        self.assertTrue(answer_run.underline)

    def test_writer_reuses_existing_answers_heading(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self.create_worksheet(directory, include_answers=True)
            parsed = DocxParser.parse(source)
            solved = SolvedWorksheet(
                qa_answers=[
                    {
                        "q_index": 1,
                        "label": "Q9",
                        "question": parsed.questions[1].text,
                        "answer": "Explained.",
                    }
                ]
            )
            output = DocxWriter(AgentConfig(output_dir=Path(directory))).update_document(
                parsed, solved, Path(directory) / "solved.docx"
            )
            result = docx.Document(output)

        headings = [
            paragraph
            for paragraph in result.paragraphs
            if paragraph.text.strip().rstrip(":").casefold() == "answers"
        ]
        self.assertEqual(len(headings), 1)
        self.assertIn("9. Explained.", [paragraph.text for paragraph in result.paragraphs])

    def test_solver_prompt_requires_paragraphs_or_bullets(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parsed = DocxParser.parse(self.create_worksheet(directory))
        prompt = GeminiSolver(AgentConfig(gemini_api_key=""))._build_prompt(parsed)

        self.assertIn("Never format an answer as a table", prompt)
        self.assertIn("Use paragraphs or simple bullet points only", prompt)
        self.assertIn("do not use a table", prompt)


if __name__ == "__main__":
    unittest.main()

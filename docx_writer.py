"""
Document updater for worksheet docx files.
Safely replaces blanks inline, injects student metadata (single-line or multi-line),
and appends structured answers at the bottom of the document when open-ended
questions exist, strictly preserving document formatting, styles, fonts, margins, headers, and footers.
"""

import logging
import re
from pathlib import Path
from typing import Dict, List, Optional
import docx
from docx.shared import Pt

from config import AgentConfig
from docx_parser import ParsedWorksheet
from gemini_solver import SolvedWorksheet

logger = logging.getLogger(__name__)


class DocxWriter:
    """Updates and formats docx documents with solutions and student metadata."""

    def __init__(self, config: AgentConfig):
        self.config = config

    def update_document(
        self,
        parsed_ws: ParsedWorksheet,
        solved_ws: SolvedWorksheet,
        output_path: Optional[Path] = None,
    ) -> Path:
        """
        Updates the docx document with student info, inline blank answers,
        and bottom QA solutions (if any), then saves it.
        """
        doc = parsed_ws.doc

        # 1. Update Student Metadata (Single or Multi line)
        self._update_student_metadata(doc, parsed_ws)

        # 2. Update Question Blanks (Inline replacement)
        for p_idx, answers in solved_ws.blank_answers.items():
            if 0 <= p_idx < len(doc.paragraphs):
                self._fill_blank_paragraph(doc.paragraphs[p_idx], answers)

        # 3. Add QA / Descriptive / Calculation Answers at the Bottom of the file
        if solved_ws.qa_answers:
            self._append_bottom_answers(doc, parsed_ws, solved_ws.qa_answers)

        # 4. Determine save destination
        if output_path is None:
            if self.config.in_place:
                output_path = parsed_ws.file_path
            else:
                self.config.output_dir.mkdir(parents=True, exist_ok=True)
                output_path = self.config.output_dir / parsed_ws.file_path.name

        doc.save(str(output_path))
        logger.info(f"Updated document saved to: {output_path}")
        return output_path

    def _update_student_metadata(self, doc: docx.Document, parsed_ws: ParsedWorksheet) -> None:
        """Updates student details for single-line or multi-line formats."""
        if parsed_ws.student_info_mode == "single" and parsed_ws.student_info_p_idx is not None:
            paragraph = doc.paragraphs[parsed_ws.student_info_p_idx]
            reg_label = (
                "Roll No.: "
                if re.search(r"\broll\s*(?:no\.?|number)\s*:", paragraph.text, re.IGNORECASE)
                else "Reg.No.: "
            )
            font_size = self._get_paragraph_font_size(paragraph)
            paragraph.text = ""

            def add_field(label: str, value: str, is_last: bool = False):
                lbl_run = paragraph.add_run(label)
                lbl_run.bold = True
                lbl_run.font.name = self.config.font_name
                if font_size:
                    lbl_run.font.size = font_size

                val_run = paragraph.add_run(value)
                val_run.bold = self.config.highlight_bold
                val_run.underline = self.config.highlight_underline
                val_run.font.name = self.config.font_name
                if font_size:
                    val_run.font.size = font_size

                if not is_last:
                    space_run = paragraph.add_run("\t\t")
                    space_run.font.name = self.config.font_name
                    if font_size:
                        space_run.font.size = font_size

            add_field("Name: ", self.config.student_name)
            add_field(reg_label, self.config.reg_no)
            add_field("Date: ", self.config.date_str, is_last=True)

        elif parsed_ws.student_info_mode == "multi":
            mapping = [
                ("name", "Name: ", self.config.student_name),
                ("reg_no", "Reg.No.: ", self.config.reg_no),
                ("date", "Date: ", self.config.date_str),
            ]
            for key, label, val in mapping:
                if key in parsed_ws.student_metadata_map:
                    p_idx = parsed_ws.student_metadata_map[key]
                    if 0 <= p_idx < len(doc.paragraphs):
                        p = doc.paragraphs[p_idx]
                        field_label = label
                        if key == "reg_no":
                            field_label = (
                                "Roll No.: "
                                if re.search(r"\broll\s*(?:no\.?|number)\s*:", p.text, re.IGNORECASE)
                                else "Reg.No.: "
                            )
                        font_size = self._get_paragraph_font_size(p)
                        p.text = ""

                        lbl_run = p.add_run(field_label)
                        lbl_run.bold = True
                        lbl_run.font.name = self.config.font_name
                        if font_size:
                            lbl_run.font.size = font_size

                        val_run = p.add_run(val)
                        val_run.bold = self.config.highlight_bold
                        val_run.underline = self.config.highlight_underline
                        val_run.font.name = self.config.font_name
                        if font_size:
                            val_run.font.size = font_size

    def _fill_blank_paragraph(
        self,
        paragraph: docx.text.paragraph.Paragraph,
        answers: List[str],
    ) -> None:
        """Replaces series of underscores with the solved answer run."""
        original_text = paragraph.text
        if not original_text.strip():
            return

        font_size = self._get_paragraph_font_size(paragraph)
        parts = re.split(r"_{3,}", original_text)

        if len(parts) == 1:
            return

        paragraph.text = ""
        for i, part in enumerate(parts):
            if part:
                p_run = paragraph.add_run(part)
                p_run.font.name = self.config.font_name
                if font_size:
                    p_run.font.size = font_size

            if i < len(parts) - 1:
                ans_val = answers[i] if i < len(answers) else ""
                ans_run = paragraph.add_run(ans_val)
                ans_run.bold = self.config.highlight_bold
                ans_run.underline = self.config.highlight_underline
                ans_run.font.name = self.config.font_name
                if font_size:
                    ans_run.font.size = font_size

    def _append_bottom_answers(
        self,
        doc: docx.Document,
        parsed_ws: ParsedWorksheet,
        qa_answers: List[Dict[str, str | int]],
    ) -> None:
        """Appends structured answers at the bottom of the file (before footer if present)."""
        footer_p = None
        if parsed_ws.footer_p_idx is not None and 0 <= parsed_ws.footer_p_idx < len(doc.paragraphs):
            footer_p = doc.paragraphs[parsed_ws.footer_p_idx]

        existing_header = next(
            (
                p
                for p in doc.paragraphs
                if p.text.strip().rstrip(":").casefold() == "answers"
            ),
            None,
        )
        if existing_header is None:
            header_p = doc.add_paragraph()
            header_p.paragraph_format.space_before = Pt(12)
            header_p.paragraph_format.space_after = Pt(6)
            header_p.paragraph_format.line_spacing = 1.15

            h_run = header_p.add_run("ANSWERS")
            h_run.bold = True
            h_run.font.name = self.config.font_name
            h_run.font.size = Pt(12)

            if footer_p is not None:
                footer_p._p.addprevious(header_p._p)

        # Add each QA answer with proper line spacing and normal font weight.
        for idx, item in enumerate(qa_answers):
            ans_p = doc.add_paragraph()
            ans_p.paragraph_format.space_before = Pt(2)
            ans_p.paragraph_format.space_after = Pt(8)
            ans_p.paragraph_format.line_spacing = 1.15

            question_label = str(item.get("label", ""))
            label_match = re.match(r"^(?:Q\s*)?(\d+)\b", question_label, re.IGNORECASE)
            if label_match:
                question_number = label_match.group(1)
            else:
                question_index = item.get("q_index")
                question_number = (
                    str(question_index + 1)
                    if isinstance(question_index, int)
                    else str(idx + 1)
                )
            label_text = f"{question_number}. "
            lbl_run = ans_p.add_run(label_text)
            lbl_run.bold = True
            lbl_run.font.name = self.config.font_name
            lbl_run.font.size = Pt(11)

            # Clean answer text: remove markdown bold/italic tags and redundant leading question prefixes
            ans_text = item.get("answer", "")
            clean_ans = re.sub(r"\*\*(.*?)\*\*", r"\1", ans_text)
            clean_ans = re.sub(r"__(.*?)__", r"\1", clean_ans)
            clean_ans = re.sub(r"^(?:(?:Q\d+|\d+|Ans|Answer)\s*[:.)-]\s*)+", "", clean_ans, flags=re.IGNORECASE).strip()

            # Answer body run - normal font weight, no bold, no underline
            ans_run = ans_p.add_run(clean_ans)
            ans_run.bold = False
            ans_run.underline = False
            ans_run.font.name = self.config.font_name
            ans_run.font.size = Pt(11)

            if footer_p is not None:
                footer_p._p.addprevious(ans_p._p)

    def _get_paragraph_font_size(self, paragraph: docx.text.paragraph.Paragraph) -> Optional[Pt]:
        """Helper to extract existing font size from runs if available."""
        for run in paragraph.runs:
            if run.font.size:
                return run.font.size
        return Pt(12)

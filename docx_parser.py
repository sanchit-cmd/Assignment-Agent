"""
Document parser for worksheet docx files.
Extracts worksheet context, student metadata locations, footer markers, and question items.
Supports both fill-in-the-blank worksheets (Unit 1) and mixed/open-ended worksheets (Unit 2).
"""

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional
import docx


@dataclass
class QuestionItem:
    question_index: int
    paragraph_indices: List[int]
    text: str
    blank_count: int
    item_type: str = "fill_in_the_blank"  # "fill_in_the_blank", "qa", "mcq"
    options: List[str] = field(default_factory=list)
    label: str = ""


@dataclass
class ParsedWorksheet:
    file_path: Path
    topic: str
    worksheet_no: str
    course_info: str
    student_info_mode: str  # "single", "multi", "none"
    student_info_p_idx: Optional[int]
    student_metadata_map: Dict[str, int]
    footer_p_idx: Optional[int]
    questions: List[QuestionItem]
    doc: docx.Document


class DocxParser:
    """Parser to extract metadata, student fields, and questions from docx worksheets."""

    BLANK_REGEX = re.compile(r"_{3,}")
    NAME_REGEX = re.compile(r"^name\s*:", re.IGNORECASE)
    REG_LABEL_PATTERN = r"(?:reg(?:istration)?\.?\s*(?:no\.?|number)|roll\s*(?:no\.?|number))"
    REG_REGEX = re.compile(rf"^{REG_LABEL_PATTERN}\s*:", re.IGNORECASE)
    REG_INLINE_REGEX = re.compile(rf"\b{REG_LABEL_PATTERN}\s*:", re.IGNORECASE)
    DATE_REGEX = re.compile(r"^date\s*:", re.IGNORECASE)
    DATE_INLINE_REGEX = re.compile(r"\bdate\s*:", re.IGNORECASE)

    SECTION_HEADER_PATTERNS = [
        r"^questions:?$",
        r"^answers:?$",
        r"^worksheet(\s+no\.?\s*[0-9.]+)?$",
        r"^total\s+marks\s*:",
        r"^\(?(mark\s+)?distribution\s*:",
        r"^instructions\s*:",
        r"^fill\s+in\s+the\s+blanks?:?$",
        r"^technical\s+fill-in-the-blanks?:?$",
        r"^technical\s+completion:?$",
        r"^scenario-based\s+(question|analysis|challenge|calculation):?$",
        r"^analytical\s+(problem|comparison):?$",
        r"^multiple\s+choice\s+question:?$",
        r"^conceptual\s+matching:?$",
        r"^short\s+answer\s+question:?$",
        r"^pattern\s+visualization:?$",
        r"^comparison\s+of\s+multiple\s+access\s+techniques\s+in\s+wireless\s+communication$",
        r"^fill\s+in\s+the\s+blanks\s+with\s+the\s+appropriate\s+.*:?$",
        r"^complete\s+the\s+following\s+statements\s+regarding\s+.*:?$",
        r"^match\s+the\s+factor\s+to\s+its\s+effect\s+.*:?$",
        r"^match\s+the\s+following\s+applications\s+to\s+.*:?$",
        r"^complete\s+the\s+following\s+comparison\s+matrix\s+.*:?$",
    ]

    INSTITUTION_KEYWORDS = [
        "srm institute",
        "school of computing",
        "department of computing",
        "college of engineering",
    ]

    @classmethod
    def is_section_header(cls, text: str) -> bool:
        t = text.strip().lower()
        for pattern in cls.SECTION_HEADER_PATTERNS:
            if re.match(pattern, t):
                return True
        return False

    @classmethod
    def is_institutional(cls, text: str) -> bool:
        t = text.strip().lower()
        return any(k in t for k in cls.INSTITUTION_KEYWORDS)

    @classmethod
    def is_course_footer(cls, text: str) -> bool:
        t = text.strip().lower()
        return bool(re.search(r"21cse458t", t))

    @classmethod
    def is_worksheet_title(cls, text: str, topic: str) -> bool:
        title = text.strip().strip("()").strip()
        return bool(topic and title.casefold() == topic.strip().casefold())

    @staticmethod
    def _question_label(text: str, fallback_index: int) -> str:
        for line in text.splitlines():
            match = re.match(r"^\s*(?:Q\s*)?(\d+)\s*[.)]?", line, re.IGNORECASE)
            if match:
                return f"Q{match.group(1)}"
        return f"Q{fallback_index}"

    @classmethod
    def is_continuation_line(
        cls,
        text: str,
        prev_text: str,
        current_blanks: int,
        current_is_mcq: bool,
    ) -> bool:
        t = text.strip()
        prev_t = prev_text.strip()
        has_blanks = bool(cls.BLANK_REGEX.search(t))

        # 1. Obvious sentence breaks without ending punctuation or hyphenated
        if prev_t.endswith("-") or prev_t.endswith(",") or (prev_t and prev_t[-1] not in ".?!:'\"" and not has_blanks):
            return True
        if t[0].islower():
            return True

        # 2. Bullet options / subparts
        if re.match(r"^(a\)|b\)|c\)|d\)|e\)|\([a-z0-9]+\)|[a-z]\.|\([A-Z]\)|§|•|-|\d+\)|[A-Z]:)", t, re.IGNORECASE):
            return True
        if re.match(r"^(characteristic|option|parameter|case)\s+[a-z0-9]+:", t, re.IGNORECASE):
            return True

        # 3. Short lines following an MCQ prompt
        if current_is_mcq and len(t) < 150 and not has_blanks and not cls.is_section_header(t):
            return True

        return False

    @classmethod
    def parse(cls, file_path: Path) -> ParsedWorksheet:
        doc = docx.Document(str(file_path))

        topic = ""
        worksheet_no = ""
        course_info = ""
        student_info_mode = "none"
        student_info_p_idx = None
        student_metadata_map: Dict[str, int] = {}
        footer_p_idx = None

        # 1. Extract course info from footers or headers if present
        for section in doc.sections:
            for p in section.footer.paragraphs:
                if p.text.strip():
                    course_info += " " + p.text.strip()
            for p in section.header.paragraphs:
                if p.text.strip():
                    course_info += " " + p.text.strip()

        # 2. Extract Topic and Worksheet No from top paragraphs
        for idx, p in enumerate(doc.paragraphs[:6]):
            text = p.text.strip()
            if not text:
                continue

            if "topic:" in text.lower() and not topic:
                topic_match = re.search(r"topic\s*:\s*(.*?)(?:\s*Worksheet|\t|\n|$)", text, re.IGNORECASE)
                if topic_match and topic_match.group(1).strip():
                    topic = topic_match.group(1).strip()

            if "worksheet" in text.lower() and not worksheet_no:
                ws_match = re.search(r"worksheet\s*no\.?\s*([0-9.]+)", text, re.IGNORECASE)
                if ws_match:
                    worksheet_no = ws_match.group(1).strip()

        # Fallback topic extraction if only in parenthesized title
        if not topic:
            for p in doc.paragraphs[:10]:
                t = p.text.strip()
                if t.startswith("(") and t.endswith(")"):
                    topic = t.strip("()")
                    break

        # 3. Detect Student Info Paragraphs (Single-line or Multi-line)
        for idx, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if not text:
                continue

            # Check combined single line
            if cls.NAME_REGEX.search(text) and (
                cls.REG_INLINE_REGEX.search(text) or cls.DATE_INLINE_REGEX.search(text)
            ):
                student_info_mode = "single"
                student_info_p_idx = idx
                break

            # Check multi-line individual fields
            if cls.NAME_REGEX.search(text):
                student_metadata_map["name"] = idx
            elif cls.REG_REGEX.search(text):
                student_metadata_map["reg_no"] = idx
            elif cls.DATE_REGEX.search(text):
                student_metadata_map["date"] = idx

        if student_info_mode == "none" and ("name" in student_metadata_map or "reg_no" in student_metadata_map):
            student_info_mode = "multi"

        # 4. Find Footer paragraph in body (e.g. "21CSE458T- WIRELESS AND MOBILE COMMUNICATION")
        for idx in range(len(doc.paragraphs) - 1, -1, -1):
            text = doc.paragraphs[idx].text.strip()
            if cls.is_course_footer(text):
                footer_p_idx = idx
                if not course_info:
                    course_info = text
                break

        # 5. Determine Question Boundaries
        q_start_idx = 0
        student_indices = [student_info_p_idx] if student_info_p_idx is not None else list(student_metadata_map.values())
        max_student_idx = max(student_indices) if student_indices else 0

        for idx, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if not text:
                continue
            if idx <= max_student_idx:
                q_start_idx = idx + 1
                continue
            if cls.is_worksheet_title(text, topic):
                q_start_idx = idx + 1
                continue
            if cls.is_institutional(text) or "topic:" in text.lower() or cls.is_section_header(text):
                q_start_idx = idx + 1
                continue
            # First actual question paragraph reached
            break

        end_idx = footer_p_idx if footer_p_idx is not None else len(doc.paragraphs)

        # 6. Group and Segment Questions
        questions: List[QuestionItem] = []
        current_paras: List[int] = []
        current_texts: List[str] = []
        current_blanks = 0
        current_is_mcq = False

        def finalize_q():
            nonlocal current_paras, current_texts, current_blanks, current_is_mcq
            if current_paras:
                full_q_text = "\n".join(current_texts)
                questions.append(
                    QuestionItem(
                        question_index=len(questions),
                        paragraph_indices=list(current_paras),
                        text=full_q_text,
                        blank_count=current_blanks,
                        item_type="fill_in_the_blank" if current_blanks > 0 else "qa",
                        label=cls._question_label(full_q_text, len(questions) + 1),
                    )
                )
                current_paras = []
                current_texts = []
                current_blanks = 0
                current_is_mcq = False

        for idx in range(q_start_idx, end_idx):
            p = doc.paragraphs[idx]
            text = p.text.strip()
            if not text:
                continue
            if cls.is_section_header(text):
                finalize_q()
                continue

            blanks = len(cls.BLANK_REGEX.findall(text))
            is_cont = False

            if current_paras:
                prev_text = current_texts[-1].strip()
                is_cont = cls.is_continuation_line(text, prev_text, current_blanks, current_is_mcq)

            if is_cont:
                current_paras.append(idx)
                current_texts.append(text)
                current_blanks += blanks
            else:
                finalize_q()
                current_paras = [idx]
                current_texts = [text]
                current_blanks = blanks
                current_is_mcq = bool(
                    re.search(r"(choose|options?|increases by|multiple choice)", text, re.IGNORECASE)
                )

        finalize_q()

        return ParsedWorksheet(
            file_path=file_path,
            topic=topic or "Wireless Communication",
            worksheet_no=worksheet_no or "Worksheet",
            course_info=course_info.strip() or "21CSE458T - Wireless and Mobile Communication",
            student_info_mode=student_info_mode,
            student_info_p_idx=student_info_p_idx,
            student_metadata_map=student_metadata_map,
            footer_p_idx=footer_p_idx,
            questions=questions,
            doc=doc,
        )

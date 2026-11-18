"""
Gemini AI Solver module for answering worksheet questions.
Uses google-genai SDK to generate accurate, course-aligned answers
for both fill-in-the-blank and open-ended/calculation/scenario questions.
"""

import json
import logging
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

from config import AgentConfig
from docx_parser import DocxParser, ParsedWorksheet, QuestionItem

logger = logging.getLogger(__name__)


class QuestionSolution(BaseModel):
    question_index: int = Field(description="0-based index of the question in the provided questions list")
    item_type: str = Field(description="'fill_in_the_blank' or 'qa'")
    blanks: List[str] = Field(
        default_factory=list,
        description="For fill_in_the_blank: list of exact strings to replace each blank in left-to-right order.",
    )
    answer_text: str = Field(
        default="",
        description="For QA / calculation / scenario / MCQ / matching: the clear, precise, syllabus-aligned solution.",
    )
    explanation: Optional[str] = Field(default="", description="Brief explanation, steps, or formula used.")


class WorksheetSolutions(BaseModel):
    answers: List[QuestionSolution]


@dataclass
class SolvedWorksheet:
    blank_answers: Dict[int, List[str]] = field(default_factory=dict)
    qa_answers: List[Dict[str, str | int]] = field(default_factory=list)


class GeminiSolver:
    """Solves worksheet questions using Google Gemini API."""

    def __init__(self, config: AgentConfig):
        self.config = config
        self.client = None
        if self.config.gemini_api_key:
            self.client = genai.Client(api_key=self.config.gemini_api_key)

    def solve_worksheet(self, parsed_ws: ParsedWorksheet) -> SolvedWorksheet:
        """
        Solves all questions in the parsed worksheet.
        Returns a SolvedWorksheet containing inline blank replacements and bottom QA solutions.
        """
        if not parsed_ws.questions:
            return SolvedWorksheet()

        if not self.client:
            raise ValueError(
                "Gemini API key is not configured. Please set GEMINI_API_KEY in .env or pass --api-key."
            )

        prompt = self._build_prompt(parsed_ws)

        for attempt in range(1, self.config.max_retries + 1):
            try:
                response = self.client.models.generate_content(
                    model=self.config.gemini_model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=self.config.temperature,
                        response_mime_type="application/json",
                        response_schema=WorksheetSolutions,
                    ),
                )

                # Parse structured JSON response
                solutions: WorksheetSolutions = WorksheetSolutions.model_validate_json(response.text)

                blank_answers: Dict[int, List[str]] = {}
                qa_answers: List[Dict[str, str | int]] = []

                for ans in solutions.answers:
                    if 0 <= ans.question_index < len(parsed_ws.questions):
                        q = parsed_ws.questions[ans.question_index]

                        if q.blank_count > 0 and ans.blanks:
                            blank_idx = 0
                            for p_idx in q.paragraph_indices:
                                p_text = parsed_ws.doc.paragraphs[p_idx].text
                                p_blanks = len(DocxParser.BLANK_REGEX.findall(p_text))
                                if p_blanks > 0:
                                    p_answers = ans.blanks[blank_idx : blank_idx + p_blanks]
                                    blank_answers[p_idx] = p_answers
                                    blank_idx += p_blanks
                        else:
                            ans_text = ans.answer_text.strip()
                            if not ans_text and ans.blanks:
                                ans_text = ", ".join(ans.blanks)
                            qa_answers.append(
                                {
                                    "q_index": ans.question_index,
                                    "label": q.label or f"Q{ans.question_index + 1}",
                                    "question": q.text,
                                    "answer": ans_text,
                                }
                            )

                return SolvedWorksheet(blank_answers=blank_answers, qa_answers=qa_answers)

            except Exception as e:
                logger.warning(
                    f"Attempt {attempt}/{self.config.max_retries} failed for {parsed_ws.file_path.name}: {e}"
                )
                if attempt < self.config.max_retries:
                    time.sleep(self.config.retry_delay * attempt)
                else:
                    raise RuntimeError(
                        f"Failed to solve worksheet {parsed_ws.file_path.name} after {self.config.max_retries} attempts: {e}"
                    ) from e

        return SolvedWorksheet()

    def _build_prompt(self, parsed_ws: ParsedWorksheet) -> str:
        """Constructs a high-accuracy, context-rich prompt for Gemini."""
        questions_text = []
        for idx, q in enumerate(parsed_ws.questions):
            q_type_hint = "Fill in the blanks" if q.blank_count > 0 else "Descriptive / Calculation / Scenario / MCQ Question"
            questions_text.append(f"[{idx}] ({q_type_hint}, {q.blank_count} blanks):\n{q.text}\n")

        prompt = f"""You are an expert professor and teaching assistant for university engineering courses, specifically in '{parsed_ws.course_info}'.

Context:
- Course: {parsed_ws.course_info}
- Topic: {parsed_ws.topic}
- Worksheet No: {parsed_ws.worksheet_no}

Questions:
{chr(10).join(questions_text)}

Instructions:
1. For each question, provide an accurate, syllabus-aligned, technical engineering solution.
2. Never format an answer as a table, including Markdown tables or aligned columns. Use paragraphs or simple bullet points only.
3. If question is 'fill_in_the_blank':
   - Provide the exact technical word or phrase in the 'blanks' list in left-to-right order.
   - Keep each blank item concise (do not repeat the whole sentence).
4. If question is 'qa' (Descriptive, Calculation, Scenario-Based, MCQ, Matching, Comparison):
   - Provide the complete, precise answer in 'answer_text'.
   - For numerical calculations, include the formula used, step-by-step substitution, and final value with appropriate units (e.g. dB, km, km², MHz).
   - For MCQs, state the selected option and brief justification.
   - For matching/comparison questions, use bullets or short paragraphs to identify each item and its correct match; do not use a table.
   - Keep answers clear, technical, and directly addressing the prompt.
"""
        return prompt

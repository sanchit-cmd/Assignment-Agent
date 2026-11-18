"""
Configuration settings for the Assignment Solver Agent.
"""

import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Load .env if present
load_dotenv()


@dataclass
class AgentConfig:
    # Student Details
    student_name: str = "Sanchit Nigam"
    reg_no: str = "RA2311003030178"
    date_str: str = datetime.now().strftime("%d/%m/%Y")

    # API Configuration
    gemini_api_key: str = os.getenv("GEMINI_API_KEY", "")
    gemini_model: str = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
    temperature: float = 0.2
    max_retries: int = 3
    retry_delay: float = 2.0

    # Paths
    input_dir: Path = Path("InputFolder")
    output_dir: Path = Path("OutputFolder")
    in_place: bool = False

    # Styling for filled answers
    highlight_underline: bool = True
    highlight_bold: bool = True
    font_name: str = "Times New Roman"

    def validate(self) -> None:
        """Validate critical configuration parameters."""
        if not self.in_place and not self.output_dir.exists():
            self.output_dir.mkdir(parents=True, exist_ok=True)

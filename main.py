"""
Assignment Solver Agent - Batch Docx Processing with Google Gemini.

Author: Sanchit Nigam (RA2311003030178)
"""

import argparse
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.progress import BarColumn, Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.table import Table

from config import AgentConfig
from docx_parser import DocxParser, ParsedWorksheet
from docx_writer import DocxWriter
from gemini_solver import GeminiSolver, SolvedWorksheet

console = Console()


def create_banner() -> Panel:
    return Panel(
        "[bold cyan]🎓 Google Gemini Assignment Solver Agent[/bold cyan]\n"
        "[dim]Batch Docx Worksheet Solver & Formatter for Engineering Assignments[/dim]\n\n"
        "👤 [bold]Student:[/bold] Sanchit Nigam | 🆔 [bold]Reg No:[/bold] RA2311003030178",
        border_style="cyan",
        title="[bold yellow]SRM IST - Wireless & Mobile Communication[/bold yellow]",
    )


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Batch process and solve docx assignment worksheets using Google Gemini API."
    )
    parser.add_argument(
        "--input-dir",
        "-i",
        type=Path,
        default=Path("InputFolder"),
        help="Path to folder containing .docx files (default: InputFolder)",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        default=Path("OutputFolder"),
        help="Path to save updated .docx files (default: OutputFolder)",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite the input docx files in-place instead of saving to output-dir",
    )
    parser.add_argument(
        "--api-key",
        type=str,
        default=os.getenv("GEMINI_API_KEY", ""),
        help="Gemini API Key (can also be set in .env as GEMINI_API_KEY)",
    )
    parser.add_argument(
        "--model",
        type=str,
        default=os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
        help="Gemini model to use (default: gemini-2.5-flash)",
    )
    parser.add_argument(
        "--name",
        type=str,
        default="Sanchit Nigam",
        help="Student name to inject (default: Sanchit Nigam)",
    )
    parser.add_argument(
        "--reg-no",
        type=str,
        default="RA2311003030178",
        help="Registration number to inject (default: RA2311003030178)",
    )
    parser.add_argument(
        "--date",
        type=str,
        default=datetime.now().strftime("%d/%m/%Y"),
        help="Date string to inject (default: current date DD/MM/YYYY)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview questions and solutions in the console without modifying files",
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        help="Run in offline mock mode (for testing workflow without using API credits)",
    )

    return parser.parse_args()


def mock_solve_worksheet(parsed_ws: ParsedWorksheet) -> SolvedWorksheet:
    """Mock answer generator for offline testing and verification."""
    blank_answers: Dict[int, List[str]] = {}
    qa_answers: List[Dict[str, str | int]] = []

    # Mock answer banks for specific Unit 2 topics
    mock_solutions = {
        # WS 2.1
        "2.1": {
            "blanks": ["cell", "hexagonal", "cell radius (R)", "hexagonal"],
            "qa": [
                "The three common shapes are equilateral triangle, square, and regular hexagon. The regular hexagon is chosen standard because it covers a given area with the minimum perimeter and fewest number of cells without any gaps or overlapping regions.",
                "Footprint is the actual geographic coverage area over which a base station provides reliable radio communication signal strength.",
                "Area of hexagon A = (3√3 / 2) * R² = 2.598 * (5)² = 64.95 km².",
                "Square area = 2 * R² = 2 * (2)² = 8 km²; Hexagon area = (3√3 / 2) * R² = 2.598 * 4 = 10.39 km². The hexagonal cell covers more area (10.39 km² vs 8 km²).",
            ],
        },
        # WS 2.2
        "2.2": {
            "blanks": ["path loss exponent", "field strength meter (or spectrum analyzer)"],
            "qa": [
                "The phenomenon is Shadowing (Large-Scale Fading) caused by obstacle attenuation (foliage absorption/scattering), which attenuates signal severely even at a short distance compared to line-of-sight propagation in open field.",
                "6 dB (Since Free Space Path Loss increases as 20 log10(d2/d1) = 20 log10(2) ≈ 6.02 dB).",
                "(i) Multipath Fading -> (B) Rapid fluctuations caused by interference of reflected waves; (ii) Shadowing -> (A) Long-term signal variation due to large obstacles.",
                "SIR is the ratio of desired carrier power to total co-channel interference power (S/I). Increasing cell radius decreases received signal power at boundary (P_r ∝ R^(-n)), which lowers SIR if interfering power remains constant.",
            ],
        },
        # WS 2.3
        "2.3": {
            "blanks": ["decreases", "decreases", "Q = D / R = √(3N)"],
            "qa": [
                "Cell radius R is the distance from the center of a hexagonal cell to any of its vertices. Hexagonal geometry is preferred because it approximates a circle with zero tiling gaps/overlaps and highest area coverage efficiency.",
                "Co-channel reuse ratio Q = D/R = √(3N), where D is the distance between co-channel cells and R is cell radius. A higher Q means larger physical separation between co-channel cells, resulting in lower co-channel interference.",
                "Cell Splitting: The process of subdividing congested cells into smaller micro-cells, each with its own low-power base station and reduced antenna height, multiplying system capacity without adding extra frequency spectrum.",
                "1. System Capacity: Increases proportionally to the number of split cells. 2. Handover Frequency: Increases significantly due to smaller cell radius and faster cell boundary crossing.",
            ],
        },
        # WS 2.4
        "2.4": {
            "blanks": [],
            "qa": [
                "Capacity is given by C = M * k * N = M * S, where cluster size N determines frequency reuse. A smaller cluster size N allows channels to be reused more frequently over the area, increasing overall system capacity, but decreases reuse distance D, raising co-channel interference.",
                "For N = 7: Number of clusters M = 28 / 7 = 4 clusters. Total system capacity C = 4 * 336 = 1,344 simultaneous channels.",
                "For N = 4: Number of clusters M = 28 / 4 = 7 clusters. Total system capacity C = 7 * 336 = 2,352 simultaneous channels.",
                "Cluster size N = 4 provides higher capacity (2,352 channels vs 1,344 channels). The trade-off is significantly increased co-channel interference (lower S/I ratio) due to smaller reuse distance.",
                "Cell Splitting scales capacity by replacing large cells with smaller micro-cells of radius R/2 and lower transmit power, allowing the same set of channel frequencies to be reused much more densely across the hotspot.",
                "Microcell Zone Concept is more effective at reducing handoffs while managing CCI because mobile stations share the same base station zone controller without executing formal base-station handovers.",
            ],
        },
    }

    # Match topic or worksheet number
    ws_key = parsed_ws.worksheet_no
    preset = mock_solutions.get(ws_key)

    blank_pool = preset["blanks"] if preset else [f"Term_{i+1}" for i in range(20)]
    qa_pool = preset["qa"] if preset else [
        f"Detailed engineering solution for {q.text[:40]}... Formula: Q = √(3N), Result = 4.58 with high spectral efficiency."
        for q in parsed_ws.questions
    ]

    blank_cursor = 0
    qa_cursor = 0

    for q in parsed_ws.questions:
        if q.blank_count > 0:
            for p_idx in q.paragraph_indices:
                p_text = parsed_ws.doc.paragraphs[p_idx].text
                p_blanks = len(DocxParser.BLANK_REGEX.findall(p_text))
                if p_blanks > 0:
                    ans_list = []
                    for _ in range(p_blanks):
                        ans_list.append(blank_pool[blank_cursor % len(blank_pool)])
                        blank_cursor += 1
                    blank_answers[p_idx] = ans_list
        else:
            qa_ans = qa_pool[qa_cursor % len(qa_pool)]
            qa_cursor += 1
            qa_answers.append(
                {
                    "q_index": q.question_index,
                    "label": q.label or f"Q{q.question_index + 1}",
                    "question": q.text,
                    "answer": qa_ans,
                }
            )

    return SolvedWorksheet(blank_answers=blank_answers, qa_answers=qa_answers)


def main():
    load_dotenv()
    console.print(create_banner())

    args = parse_arguments()

    config = AgentConfig(
        student_name=args.name,
        reg_no=args.reg_no,
        date_str=args.date,
        gemini_api_key=args.api_key,
        gemini_model=args.model,
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        in_place=args.in_place,
    )
    config.validate()

    # Discover all .docx files in input directory
    if not config.input_dir.exists():
        console.print(f"[bold red]Error:[/bold red] Input directory '{config.input_dir}' does not exist!")
        sys.exit(1)

    docx_files = sorted(
        [
            f for f in config.input_dir.glob("*.docx")
            if not f.name.startswith("~$")
        ]
    )

    if not docx_files:
        console.print(f"[bold yellow]Warning:[/bold yellow] No .docx files found in '{config.input_dir}'!")
        sys.exit(0)

    console.print(
        f"📂 Found [bold cyan]{len(docx_files)}[/bold cyan] docx worksheets to process in [bold]{config.input_dir}[/bold].\n"
    )

    # Check API key if not in mock mode
    if not args.mock and not config.gemini_api_key:
        console.print(
            Panel(
                "[bold red]GEMINI_API_KEY Not Found![/bold red]\n\n"
                "To solve worksheets using Gemini AI, please:\n"
                "1. Create a [bold].env[/bold] file with: [green]GEMINI_API_KEY=your_key_here[/green]\n"
                "2. Or pass your key via CLI: [green]--api-key your_key_here[/green]\n"
                "3. Or test offline using: [green]--mock[/green]\n",
                border_style="red",
            )
        )
        sys.exit(1)

    solver = GeminiSolver(config) if not args.mock else None
    writer = DocxWriter(config)

    summary_rows = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task_id = progress.add_task("[cyan]Processing worksheets...", total=len(docx_files))

        for file_path in docx_files:
            progress.update(task_id, description=f"[cyan]Processing {file_path.name}...")
            start_t = time.time()
            try:
                # 1. Parse document
                parsed_ws = DocxParser.parse(file_path)

                # 2. Solve questions
                if args.mock:
                    solved_ws = mock_solve_worksheet(parsed_ws)
                else:
                    solved_ws = solver.solve_worksheet(parsed_ws)

                # 3. Dry run or update document
                if args.dry_run:
                    status = "[yellow]Dry Run[/yellow]"
                else:
                    out_path = writer.update_document(parsed_ws, solved_ws)
                    status = "[bold green]✓ Done[/bold green]"

                elapsed = time.time() - start_t
                total_blanks = sum(len(b) for b in solved_ws.blank_answers.values())
                total_qa = len(solved_ws.qa_answers)
                q_summary = f"{len(parsed_ws.questions)} ({total_blanks}B, {total_qa}QA)"

                summary_rows.append(
                    (
                        file_path.name,
                        parsed_ws.worksheet_no,
                        parsed_ws.topic[:35],
                        q_summary,
                        status,
                        f"{elapsed:.2f}s",
                    )
                )

            except Exception as e:
                elapsed = time.time() - start_t
                console.print(f"[red]Error processing {file_path.name}: {e}[/red]")
                summary_rows.append(
                    (
                        file_path.name,
                        "Unknown",
                        "Error",
                        "0",
                        f"[bold red]✗ Failed: {e}[/bold red]",
                        f"{elapsed:.2f}s",
                    )
                )
            finally:
                progress.advance(task_id)

    # Print Summary Table
    table = Table(title="\n📊 Batch Processing Summary Report", border_style="cyan")
    table.add_column("File Name", style="bold white")
    table.add_column("Worksheet", style="cyan")
    table.add_column("Topic", style="magenta")
    table.add_column("Questions", justify="right", style="green")
    table.add_column("Status", justify="center")
    table.add_column("Time", justify="right", style="dim")

    for row in summary_rows:
        table.add_row(*row)

    console.print(table)
    if not args.dry_run:
        dest = config.input_dir if config.in_place else config.output_dir
        console.print(
            f"\n🎉 [bold green]Batch processing completed successfully![/bold green] All updated documents are in: [bold cyan]{dest}[/bold cyan]\n"
        )


if __name__ == "__main__":
    main()

from __future__ import annotations

from pathlib import Path


def gather_pdfs(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".pdf":
            raise ValueError(f"Input file is not a PDF: {input_path}")
        return [input_path]
    if input_path.is_dir():
        return sorted(input_path.rglob("*.pdf"))
    raise ValueError(f"Input path does not exist: {input_path}")


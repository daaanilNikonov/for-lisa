#!/usr/bin/env python3
"""Сборка PDF-резюме Попова С. Д. из HTML."""

from pathlib import Path

from weasyprint import HTML

ROOT = Path(__file__).resolve().parents[1]
RESUME_DIR = ROOT / "resume"
SRC = RESUME_DIR / "index.html"
OUT = RESUME_DIR / "Попов_Семён_Backend_разработчик.pdf"


def main() -> None:
    HTML(filename=str(SRC), base_url=str(RESUME_DIR)).write_pdf(str(OUT))
    print(f"Written {OUT}")


if __name__ == "__main__":
    main()

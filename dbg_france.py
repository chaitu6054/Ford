"""One-off diagnostic: run the France stage with a full traceback.

Usage (from the project root):
    python dbg_france.py
"""
import sys
import traceback
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "astro" / "include"))

from remarketing.france_reports import generate_france_reports, has_france_sheets


def main() -> None:
    try:
        src = next((ROOT / "local" / "input").glob("*.xlsx"))
        print("input:", src.name)
        print("has_france_sheets:", has_france_sheets(src))
        result = generate_france_reports(
            input_path=src,
            output_dir=ROOT / "local" / "output" / "dbg",
            processing_date=date(2026, 9, 21),
        )
        print("OK")
        print("ERA:", result.era_path)
        print("Portefeuille:", result.portefeuille_path)
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    main()

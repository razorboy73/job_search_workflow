r"""
Template for a WAT tool. Copy this file, rename it, and fill in main().

Conventions:
- One tool, one job. Inputs come from CLI args; secrets come from .env.
- Write intermediate files to .tmp/ (disposable).
- Print a short result to stdout; exit non-zero on failure so the agent notices.

Run with:  .venv\Scripts\python.exe tools\<name>.py --help
"""

import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
TMP = ROOT / ".tmp"
load_dotenv(ROOT / ".env")


def main() -> int:
    parser = argparse.ArgumentParser(description="Describe what this tool does.")
    # parser.add_argument("--input", required=True, help="...")
    args = parser.parse_args()

    TMP.mkdir(exist_ok=True)
    # api_key = os.environ["SOME_API_KEY"]
    # ... do the work ...
    print("ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())

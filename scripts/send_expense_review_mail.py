"""CLI entry: python scripts/send_expense_review_mail.py --dry-run"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from finance_tools.expense_review_mail import main


if __name__ == "__main__":
    raise SystemExit(main())

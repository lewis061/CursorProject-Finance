from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
LOGS = ROOT / "logs"
SCRIPTS = ROOT / "scripts"

# Shared network folders for monthly expense-review mail
EMAIL_ZONE = Path(r"R:\費用檢討E-MAIL專區")
RECIPIENT_DIR = EMAIL_ZONE / "E-MIAL"
ATTACHMENT_DIR = EMAIL_ZONE / "夾檔"
CONTENT_FILE = EMAIL_ZONE / "MAIL內文.xlsx"


def ensure_dirs() -> None:
    for p in (DATA, OUTPUT, LOGS):
        p.mkdir(parents=True, exist_ok=True)

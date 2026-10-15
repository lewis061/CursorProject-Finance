from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
LOGS = ROOT / "logs"
SCRIPTS = ROOT / "scripts"


def ensure_dirs() -> None:
    for p in (DATA, OUTPUT, LOGS):
        p.mkdir(parents=True, exist_ok=True)

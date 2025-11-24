from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE = BASE_DIR / "templates" / "template.xlsx"
DEFAULT_OUTPUT = BASE_DIR / "output" / "tasks.xlsx"

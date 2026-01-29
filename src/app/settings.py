from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE = BASE_DIR / "templates" / "template.xlsx"
DEFAULT_OUTPUT = BASE_DIR / "output" / "tasks.xlsx"

# PlantUML settings
PLANTUML_SERVER_URL = "http://localhost:8080"
PLANTUML_TEMPLATE = BASE_DIR / "templates" / "plantuml_template.xlsx"
PLANTUML_OUTPUT = BASE_DIR / "output" / "plantuml_diagram.xlsx"

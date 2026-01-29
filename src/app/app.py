from __future__ import annotations

from pathlib import Path

from fastmcp import FastMCP

from ..helpers.excel import write_task_lists_to_excel
from ..helpers.plantuml import (
    generate_plantuml_image,
    write_plantuml_to_excel,
    write_plantuml_image_only,
)


# BASE_DIR = Path(__file__).resolve().parents[2]
# DEFAULT_TEMPLATE = BASE_DIR / "templates" / "template.xlsx"
# DEFAULT_OUTPUT = BASE_DIR / "output" / "tasks.xlsx"

from . import settings


server = FastMCP(
    name="excel-task-server",
    instructions=(
        "Generate numbered main/support tasks and write them into the Excel template "
        "using marker cells (marker[main_task], marker[support_task]). "
        "Use the write_task_list tool to fill templates/template.xlsx and preserve formatting. "
        "Also supports generating PlantUML diagrams and embedding them in Excel sheets. "
        "Use write_plantuml_to_excel to create diagrams from PlantUML code."
    ),
)


@server.tool(name="write_task_list")
def write_task_list(
    main_tasks: list[str],
    support_tasks: list[str],
    template_path: str | None = None,
    output_path: str | None = None,
) -> str:
    """Write numbered main and support tasks into the Excel template using marker cells."""
    template = Path(template_path) if template_path else settings.DEFAULT_TEMPLATE
    output = Path(output_path) if output_path else settings.DEFAULT_OUTPUT

    saved_path = write_task_lists_to_excel(
        template_path=template,
        output_path=output,
        main_tasks=main_tasks,
        support_tasks=support_tasks,
    )

    return (
        f"Wrote {len(main_tasks)} main tasks and {len(support_tasks)} support tasks "
        f"to {saved_path}"
    )


@server.prompt(name="random_task_list")
def random_task_list_prompt(
    main_count: int = 3,
    support_count: int = 3,
    template_path: str | None = None,
    output_path: str | None = None,
):
    """Prompt that asks the model to invent tasks and save them via the MCP tool."""
    template = template_path or str(settings.DEFAULT_TEMPLATE)
    output = output_path or str(settings.DEFAULT_OUTPUT)
    return [
        {
            "role": "system",
            "content": (
                "Create concise, varied project tasks. Use two sections: main tasks and "
                "support tasks. Each item must be numbered like '1. ...', '2. ...'. "
                "After producing the lists, call the tool `write_task_list` to save them "
                "into the Excel template while keeping existing formatting."
            ),
        },
        {
            "role": "user",
            "content": (
                f"Generate {main_count} main tasks and {support_count} support tasks. "
                f"Save them with `write_task_list`, using template_path='{template}' "
                f"and output_path='{output}'. Keep tasks under ten words. "
                "Confirm where you wrote the file after the tool call."
            ),
        },
    ]


# =============================================================================
# PlantUML Tools
# =============================================================================


@server.tool(name="write_plantuml_to_excel")
def write_plantuml_to_excel_tool(
    puml_code: str,
    template_path: str | None = None,
    output_path: str | None = None,
    server_url: str | None = None,
    image_width: int | None = None,
    image_height: int | None = None,
) -> str:
    """Write PlantUML code and its rendered image to an Excel file.
    
    Uses marker cells in the template:
    - marker[plantuml]: PlantUML source code will be written below this cell
    - marker[plantuml_image]: The generated image will be placed at this cell
    
    Args:
        puml_code: The PlantUML diagram code (with or without @startuml/@enduml)
        template_path: Path to the Excel template (default: templates/plantuml_template.xlsx)
        output_path: Path for the output file (default: output/plantuml_diagram.xlsx)
        server_url: PlantUML server URL (default: http://localhost:8080)
        image_width: Optional width for the image in pixels
        image_height: Optional height for the image in pixels
    """
    template = Path(template_path) if template_path else settings.PLANTUML_TEMPLATE
    output = Path(output_path) if output_path else settings.PLANTUML_OUTPUT
    url = server_url or settings.PLANTUML_SERVER_URL
    
    saved_path = write_plantuml_to_excel(
        template_path=template,
        output_path=output,
        puml_code=puml_code,
        server_url=url,
        image_width=image_width,
        image_height=image_height,
    )
    
    return f"Successfully wrote PlantUML diagram to {saved_path}"


@server.tool(name="write_plantuml_image")
def write_plantuml_image_tool(
    puml_code: str,
    excel_path: str,
    output_path: str | None = None,
    cell_anchor: str = "A1",
    sheet_name: str | None = None,
    server_url: str | None = None,
    image_width: int | None = None,
    image_height: int | None = None,
) -> str:
    """Write a PlantUML image to a specific cell in an Excel file.
    
    This is a simpler version that doesn't require marker cells.
    
    Args:
        puml_code: The PlantUML diagram code
        excel_path: Path to the source Excel file
        output_path: Path for the output file (default: same as excel_path with _puml suffix)
        cell_anchor: Cell reference where the image will be placed (e.g., "B5")
        sheet_name: Name of the sheet to write to (uses active sheet if None)
        server_url: PlantUML server URL (default: http://localhost:8080)
        image_width: Optional width for the image in pixels
        image_height: Optional height for the image in pixels
    """
    source = Path(excel_path)
    if output_path:
        output = Path(output_path)
    else:
        output = source.parent / f"{source.stem}_puml{source.suffix}"
    
    url = server_url or settings.PLANTUML_SERVER_URL
    
    saved_path = write_plantuml_image_only(
        excel_path=source,
        output_path=output,
        puml_code=puml_code,
        server_url=url,
        cell_anchor=cell_anchor,
        sheet_name=sheet_name,
        image_width=image_width,
        image_height=image_height,
    )
    
    return f"Successfully wrote PlantUML image to cell {cell_anchor} in {saved_path}"


@server.tool(name="generate_plantuml_png")
def generate_plantuml_png_tool(
    puml_code: str,
    output_path: str | None = None,
    server_url: str | None = None,
) -> str:
    """Generate a PlantUML diagram as a PNG file.
    
    Args:
        puml_code: The PlantUML diagram code
        output_path: Path for the output PNG file (default: output/diagram.png)
        server_url: PlantUML server URL (default: http://localhost:8080)
    """
    url = server_url or settings.PLANTUML_SERVER_URL
    output = Path(output_path) if output_path else settings.BASE_DIR / "output" / "diagram.png"
    
    image_data = generate_plantuml_image(puml_code, server_url=url, output_format="png")
    
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(image_data)
    
    return f"Successfully generated PlantUML diagram at {output}"


# =============================================================================
# PlantUML Prompts
# =============================================================================


@server.prompt(name="draw_plantuml")
def draw_plantuml_prompt(
    diagram_description: str,
    diagram_type: str = "sequence",
    template_path: str | None = None,
    output_path: str | None = None,
):
    """Prompt that asks the model to create a PlantUML diagram and save it to Excel."""
    template = template_path or str(settings.PLANTUML_TEMPLATE)
    output = output_path or str(settings.PLANTUML_OUTPUT)
    
    diagram_examples = {
        "sequence": """@startuml
actor User
participant "Web Server" as Server
database "Database" as DB

User -> Server: HTTP Request
Server -> DB: Query
DB --> Server: Results
Server --> User: HTTP Response
@enduml""",
        "class": """@startuml
class Animal {
  +name: String
  +age: int
  +eat()
  +sleep()
}

class Dog {
  +breed: String
  +bark()
}

class Cat {
  +color: String
  +meow()
}

Animal <|-- Dog
Animal <|-- Cat
@enduml""",
        "usecase": """@startuml
left to right direction
actor Customer
actor Admin

rectangle "E-Commerce System" {
  Customer --> (Browse Products)
  Customer --> (Add to Cart)
  Customer --> (Checkout)
  Admin --> (Manage Products)
  Admin --> (View Orders)
}
@enduml""",
        "activity": """@startuml
start
:Initialize;
if (Condition?) then (yes)
  :Process A;
else (no)
  :Process B;
endif
:Finalize;
stop
@enduml""",
        "state": """@startuml
[*] --> Idle
Idle --> Processing : start
Processing --> Completed : success
Processing --> Failed : error
Completed --> [*]
Failed --> Idle : retry
@enduml""",
    }
    
    example = diagram_examples.get(diagram_type, diagram_examples["sequence"])
    
    return [
        {
            "role": "system",
            "content": (
                f"You are an expert at creating PlantUML diagrams. Create a {diagram_type} diagram "
                "based on the user's description. Use proper PlantUML syntax with @startuml and @enduml tags. "
                "After creating the diagram code, call the `write_plantuml_to_excel` tool to save it.\n\n"
                f"Example {diagram_type} diagram:\n```plantuml\n{example}\n```"
            ),
        },
        {
            "role": "user",
            "content": (
                f"Create a PlantUML {diagram_type} diagram for: {diagram_description}\n\n"
                f"After creating the PlantUML code, call `write_plantuml_to_excel` with:\n"
                f"- puml_code: Your PlantUML code\n"
                f"- template_path: '{template}'\n"
                f"- output_path: '{output}'\n"
                f"- server_url: '{settings.PLANTUML_SERVER_URL}'\n\n"
                "Confirm the output path after saving."
            ),
        },
    ]


@server.prompt(name="plantuml_from_code")
def plantuml_from_code_prompt(
    source_code: str,
    diagram_type: str = "class",
    output_path: str | None = None,
):
    """Prompt that asks the model to analyze source code and generate a PlantUML diagram."""
    output = output_path or str(settings.PLANTUML_OUTPUT)
    
    return [
        {
            "role": "system",
            "content": (
                "You are an expert at analyzing source code and creating PlantUML diagrams. "
                f"Analyze the provided code and create a {diagram_type} diagram that represents "
                "its structure, relationships, or flow. Use proper PlantUML syntax."
            ),
        },
        {
            "role": "user",
            "content": (
                f"Analyze this code and create a PlantUML {diagram_type} diagram:\n\n"
                f"```\n{source_code}\n```\n\n"
                f"After creating the PlantUML code, call `write_plantuml_to_excel` with:\n"
                f"- puml_code: Your PlantUML code\n"
                f"- output_path: '{output}'\n"
                f"- server_url: '{settings.PLANTUML_SERVER_URL}'\n\n"
                "Explain what the diagram shows and confirm the output path."
            ),
        },
    ]

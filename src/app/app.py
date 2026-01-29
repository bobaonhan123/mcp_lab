from __future__ import annotations

from pathlib import Path

from fastmcp import FastMCP

from ..helpers.excel import write_task_lists_to_excel
from ..helpers.plantuml import (
    generate_plantuml_image,
    write_plantuml_to_excel,
    write_plantuml_image_only,
)
from ..helpers.code import (
    capture_code_block,
    capture_multiple_blocks,
    format_code_block,
    create_code_documentation,
)
from ..helpers.code_excel import (
    write_code_blocks_to_excel,
    write_code_and_diagram_to_excel,
    capture_and_write_to_excel,
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
        "Use write_plantuml_to_excel to create diagrams from PlantUML code. "
        "Also supports capturing code blocks from source files with line numbers. "
        "Use capture_code_block to extract code from files."
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


# =============================================================================
# Code Capture Tools
# =============================================================================


@server.tool(name="capture_code_block")
def capture_code_block_tool(
    file_path: str,
    start_line: int,
    end_line: int | None = None,
) -> str:
    """Capture a code block from a source file with line numbers.
    
    Args:
        file_path: Path to the source file (relative or absolute)
        start_line: Starting line number (1-indexed)
        end_line: Ending line number (1-indexed, inclusive). If None, captures single line.
        
    Returns:
        Formatted code block with file path, line numbers, and syntax highlighting
    """
    try:
        block = capture_code_block(file_path, start_line, end_line)
        return format_code_block(block, include_line_numbers=True)
    except Exception as e:
        return f"Error capturing code block: {str(e)}"


@server.tool(name="capture_multiple_blocks")
def capture_multiple_blocks_tool(
    file_path: str,
    ranges: list[list[int]],
) -> str:
    """Capture multiple code blocks from the same file.
    
    Args:
        file_path: Path to the source file
        ranges: List of [start_line, end_line] pairs (1-indexed)
        
    Returns:
        Formatted code documentation with all captured blocks
    """
    try:
        # Convert list of lists to list of tuples
        range_tuples = [(r[0], r[1] if len(r) > 1 else r[0]) for r in ranges]
        blocks = capture_multiple_blocks(file_path, range_tuples)
        doc = create_code_documentation(blocks, title=f"Code Blocks from {file_path}")
        return doc
    except Exception as e:
        return f"Error capturing code blocks: {str(e)}"


# =============================================================================
# Code Capture Prompts
# =============================================================================


@server.prompt(name="capture_code")
def capture_code_prompt(
    file_path: str,
    description: str = "Review and document this code",
):
    """Prompt that asks the model to capture and document code blocks."""
    return [
        {
            "role": "system",
            "content": (
                "You are an expert code reviewer and documenter. "
                "Your task is to capture relevant code blocks from the specified file and provide documentation. "
                "Use the capture_code_block tool to extract code with line numbers. "
                "After capturing, explain what the code does, its purpose, and any important details."
            ),
        },
        {
            "role": "user",
            "content": (
                f"File: {file_path}\n\n"
                f"Task: {description}\n\n"
                "Please:\n"
                "1. Examine the file and identify key code sections\n"
                "2. Use `capture_code_block` to extract relevant sections with line numbers\n"
                "3. Explain what each captured block does\n"
                "4. Provide any relevant documentation or insights"
            ),
        },
    ]


@server.prompt(name="analyze_code_structure")
def analyze_code_structure_prompt(
    file_path: str,
):
    """Prompt that asks the model to analyze and document the overall code structure."""
    return [
        {
            "role": "system",
            "content": (
                "You are an expert software architect and code analyzer. "
                "Your task is to analyze the structure of a code file, identify key components, "
                "and document them with specific line ranges. "
                "Use capture_code_block to extract relevant sections."
            ),
        },
        {
            "role": "user",
            "content": (
                f"Analyze the code structure of: {file_path}\n\n"
                "Please:\n"
                "1. Identify the main components (classes, functions, modules, etc.)\n"
                "2. For each component, use `capture_code_block` to show the relevant code with line numbers\n"
                "3. Explain the relationships between components\n"
                "4. Describe the overall architecture and design patterns used\n"
                "5. Suggest any improvements or observations"
            ),
        },
    ]


# =============================================================================
# Code to Excel Tools
# =============================================================================


@server.tool(name="write_code_to_excel")
def write_code_to_excel_tool(
    file_path: str,
    ranges: list[list[int]],
    output_path: str | None = None,
    title: str = "Code Documentation",
) -> str:
    """Capture code blocks from a file and write to Excel.
    
    Args:
        file_path: Path to the source code file
        ranges: List of [start_line, end_line] pairs (1-indexed)
        output_path: Path for the output Excel file (default: output/code_blocks.xlsx)
        title: Title for the document
        
    Returns:
        Confirmation message with output path
    """
    try:
        output = Path(output_path) if output_path else settings.BASE_DIR / "output" / "code_blocks.xlsx"
        range_tuples = [(r[0], r[1] if len(r) > 1 else r[0]) for r in ranges]
        saved_path = capture_and_write_to_excel(file_path, range_tuples, output, title=title)
        return f"Successfully wrote {len(ranges)} code blocks to {saved_path}"
    except Exception as e:
        return f"Error writing code to Excel: {str(e)}"


@server.tool(name="write_code_and_diagram_to_excel")
def write_code_and_diagram_to_excel_tool(
    file_path: str,
    ranges: list[list[int]],
    puml_code: str,
    output_path: str | None = None,
    server_url: str | None = None,
) -> str:
    """Capture code blocks and create PlantUML diagram in Excel with 2 sheets.
    
    Creates an Excel file with:
    - Sheet 1 "Code Blocks": The captured code with line numbers
    - Sheet 2 "Sequence Diagram": The PlantUML code and rendered diagram
    
    Args:
        file_path: Path to the source code file
        ranges: List of [start_line, end_line] pairs (1-indexed)
        puml_code: PlantUML diagram code
        output_path: Path for the output Excel file (default: output/code_with_diagram.xlsx)
        server_url: PlantUML server URL (default: http://localhost:8080)
        
    Returns:
        Confirmation message with output path
    """
    try:
        output = Path(output_path) if output_path else settings.BASE_DIR / "output" / "code_with_diagram.xlsx"
        url = server_url or settings.PLANTUML_SERVER_URL
        
        # Capture code blocks
        range_tuples = [(r[0], r[1] if len(r) > 1 else r[0]) for r in ranges]
        blocks = capture_multiple_blocks(file_path, range_tuples)
        
        # Write both to Excel
        saved_path = write_code_and_diagram_to_excel(
            code_blocks=blocks,
            puml_code=puml_code,
            output_path=output,
            server_url=url,
        )
        
        return f"Successfully created Excel with {len(blocks)} code blocks and sequence diagram at {saved_path}"
    except Exception as e:
        return f"Error: {str(e)}"


# =============================================================================
# Code to Excel Prompts
# =============================================================================


@server.prompt(name="capture_code_to_excel")
def capture_code_to_excel_prompt(
    file_path: str,
    description: str = "Document this code",
):
    """Prompt that asks the model to capture code blocks and save to Excel."""
    output = str(settings.BASE_DIR / "output" / "code_blocks.xlsx")
    
    return [
        {
            "role": "system",
            "content": (
                "You are an expert code documenter. Your task is to:\n"
                "1. Analyze the code file and identify important sections\n"
                "2. Determine the line ranges for key functions, classes, or code blocks\n"
                "3. Use the `write_code_to_excel` tool to save them to an Excel file\n\n"
                "The Excel file will have formatted code with line numbers."
            ),
        },
        {
            "role": "user",
            "content": (
                f"File: {file_path}\n"
                f"Task: {description}\n\n"
                "Please:\n"
                "1. Identify the important code sections in this file\n"
                "2. Determine the line ranges (start_line, end_line) for each section\n"
                "3. Call `write_code_to_excel` with:\n"
                f"   - file_path: '{file_path}'\n"
                "   - ranges: [[start1, end1], [start2, end2], ...]\n"
                f"   - output_path: '{output}'\n"
                "4. Confirm what was saved"
            ),
        },
    ]


@server.prompt(name="document_code_with_diagram")
def document_code_with_diagram_prompt(
    file_path: str,
    diagram_description: str = "Create a sequence diagram showing the code flow",
):
    """Prompt that asks the model to create Excel with code blocks and sequence diagram."""
    output = str(settings.BASE_DIR / "output" / "code_with_diagram.xlsx")
    
    return [
        {
            "role": "system",
            "content": (
                "You are an expert code documenter and diagram creator. Your task is to:\n"
                "1. Analyze the code file and identify important sections\n"
                "2. Determine the line ranges for key functions or methods\n"
                "3. Create a PlantUML sequence diagram that shows how the code works\n"
                "4. Use the `write_code_and_diagram_to_excel` tool to save both to Excel\n\n"
                "The Excel file will have 2 sheets:\n"
                "- Sheet 1: Code blocks with line numbers\n"
                "- Sheet 2: Sequence diagram (code and rendered image)"
            ),
        },
        {
            "role": "user",
            "content": (
                f"File: {file_path}\n"
                f"Diagram: {diagram_description}\n\n"
                "Please:\n"
                "1. Analyze the code and identify important sections with their line ranges\n"
                "2. Create a PlantUML sequence diagram (use @startuml/@enduml)\n"
                "3. Call `write_code_and_diagram_to_excel` with:\n"
                f"   - file_path: '{file_path}'\n"
                "   - ranges: [[start1, end1], [start2, end2], ...]\n"
                "   - puml_code: 'Your PlantUML code'\n"
                f"   - output_path: '{output}'\n"
                "4. Explain what was captured and what the diagram shows"
            ),
        },
    ]

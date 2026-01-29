# Copilot MCP Excel Task Server

This project exposes a simple MCP server that writes numbered main/support tasks into an Excel template using marker cells (`marker[main_task]` and `marker[support_task]`). It also supports generating PlantUML diagrams and embedding them in Excel files. A Copilot-ready prompt is included to auto-generate tasks and call the MCP tool.

## Prerequisites
- Python 3.11+ recommended
- Windows shell examples use `powershell`; adjust paths for other shells
- Virtual environment (recommended): `python -m venv venv`
- **For PlantUML**: Docker installed and running

## Install dependencies
```powershell
.\venv\Scripts\Activate
pip install -r requirements.txt
```

## Run the PlantUML Server (Docker)
Before using PlantUML features, start the PlantUML server:
```powershell
docker run -d -p 8080:8080 plantuml/plantuml-server:jetty
```
This runs the PlantUML server on http://localhost:8080.

## Run the MCP server
From the repo root:
```powershell
.\venv\Scripts\python -m src.app
```
This starts the FastMCP server defined in `src/app/app.py`.

## Default file locations
- Template: `templates/template.xlsx` (contains marker cells for tasks)
- PlantUML Template: `templates/plantuml_template.xlsx` (contains marker cells for diagrams)
- Output: `output/tasks.xlsx` (created/overwritten by the tool)
- PlantUML Output: `output/plantuml_diagram.xlsx`

## MCP Tools

### Task Tools
- **`write_task_list(main_tasks, support_tasks, template_path?, output_path?)`**
  - Writes numbered tasks under the marker columns, preserving the row style beneath each marker.

### PlantUML Tools
- **`write_plantuml_to_excel(puml_code, template_path?, output_path?, server_url?, image_width?, image_height?)`**
  - Writes PlantUML code and rendered image to Excel using marker cells (`marker[plantuml]`, `marker[plantuml_image]`)
  
- **`write_plantuml_image(puml_code, excel_path, output_path?, cell_anchor?, sheet_name?, server_url?, image_width?, image_height?)`**
  - Writes a PlantUML image to a specific cell in any Excel file (no markers required)
  
- **`generate_plantuml_png(puml_code, output_path?, server_url?)`**
  - Generates a PlantUML diagram as a standalone PNG file

### Code Capture Tools
- **`capture_code_block(file_path, start_line, end_line?)`**
  - Captures a code block from a source file with line numbers (1-indexed)
  - Returns formatted code with file path, line numbers, and syntax highlighting
  - `end_line` is optional; if omitted, captures single line
  
- **`capture_multiple_blocks(file_path, ranges)`**
  - Captures multiple code blocks from the same file
  - `ranges`: List of [start_line, end_line] pairs (1-indexed)
  - Returns formatted documentation with all captured blocks

### Code to Excel Tools
- **`write_code_to_excel(file_path, ranges, output_path?, title?)`**
  - Captures code blocks and writes to Excel with formatted line numbers
  - Creates a single sheet with all code blocks
  
- **`write_code_and_diagram_to_excel(file_path, ranges, puml_code, output_path?, server_url?)`**
  - Creates Excel with **2 sheets**:
    - Sheet 1 "Code Blocks": Captured code with line numbers
    - Sheet 2 "Sequence Diagram": PlantUML code and rendered diagram image

## MCP Prompts

### Task Prompts
- **`random_task_list`** - Asks the model to invent tasks and call `write_task_list`.

### PlantUML Prompts
- **`draw_plantuml`** - Asks the model to create a PlantUML diagram based on a description.
  - Parameters: `diagram_description`, `diagram_type` (sequence/class/usecase/activity/state)
  
- **`plantuml_from_code`** - Asks the model to analyze source code and generate a PlantUML diagram.
  - Parameters: `source_code`, `diagram_type`

### Code Capture Prompts
- **`capture_code`** - Asks the model to capture and document code blocks from a file.
  - Parameters: `file_path`, `description` (optional)
  - Extracts relevant code sections with line numbers
  
- **`analyze_code_structure`** - Asks the model to analyze the overall code structure and architecture.
  - Parameters: `file_path`
  - Identifies components, relationships, and design patterns

### Code to Excel Prompts
- **`capture_code_to_excel`** - Captures code blocks and saves to Excel file.
  - Parameters: `file_path`, `description` (optional)
  - Creates formatted Excel with line numbers
  
- **`document_code_with_diagram`** - Creates Excel with code blocks AND sequence diagram (2 sheets).
  - Parameters: `file_path`, `diagram_description` (optional)
  - Sheet 1: Code blocks with line numbers
  - Sheet 2: PlantUML sequence diagram with rendered image

## Using with GitHub Copilot / MCP client
1. Start the PlantUML server: `docker run -d -p 8080:8080 plantuml/plantuml-server:jetty`
2. Start the MCP server (`python -m src.app`).
3. In your MCP-enabled client (e.g., Copilot Chat), select the server and choose a prompt:
   - `random_task_list` for generating tasks
   - `draw_plantuml` for creating diagrams
4. The model will generate content and invoke the appropriate tool.

## Customizing
- Change template/output paths by passing `template_path`/`output_path` to tools or by editing the prompt text.
- Task counts can be adjusted via the prompt arguments (`main_count`, `support_count`).
- PlantUML server URL can be changed via `server_url` parameter (default: http://localhost:8080).

## Prompt Files
- `prompts/copilot_random_task_prompt.txt` - Ready-to-paste task generation prompt
- `prompts/copilot_plantuml_prompt.txt` - Ready-to-paste PlantUML diagram generation prompt
- `prompts/copilot_capture_code_prompt.txt` - Ready-to-paste code capture prompt
- `prompts/copilot_document_with_diagram_prompt.txt` - Ready-to-paste code + diagram to Excel prompt

## Examples

### Capturing a code block
```
User: Use the capture_code prompt on src/helpers/code.py to document the capture_code_block function
AI: Captures lines with the function definition and explains its purpose
```

### Analyzing code structure
```
User: Use analyze_code_structure prompt on src/app/app.py
AI: Identifies all tools and prompts, captures relevant sections, explains the architecture
```

### Creating diagrams
```
User: Use draw_plantuml to create a sequence diagram showing how code capture works
AI: Generates PlantUML diagram and saves it to Excel
```

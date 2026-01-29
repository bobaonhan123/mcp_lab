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

## MCP Prompts

### Task Prompts
- **`random_task_list`** - Asks the model to invent tasks and call `write_task_list`.

### PlantUML Prompts
- **`draw_plantuml`** - Asks the model to create a PlantUML diagram based on a description.
  - Parameters: `diagram_description`, `diagram_type` (sequence/class/usecase/activity/state)
  
- **`plantuml_from_code`** - Asks the model to analyze source code and generate a PlantUML diagram.
  - Parameters: `source_code`, `diagram_type`

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

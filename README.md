# Copilot MCP Excel Task Server

This project exposes a simple MCP server that writes numbered main/support tasks into an Excel template using marker cells (`marker[main_task]` and `marker[support_task]`). A Copilot-ready prompt is included to auto-generate tasks and call the MCP tool.

## Prerequisites
- Python 3.11+ recommended
- Windows shell examples use `powershell`; adjust paths for other shells
- Virtual environment (recommended): `python -m venv venv`

## Install dependencies
```powershell
.\venv\Scripts\Activate
pip install -r requirements.txt
```

## Run the MCP server
From the repo root:
```powershell
.\venv\Scripts\python -m src.app
```
This starts the FastMCP server defined in `src/app/app.py`.

## Default file locations
- Template: `templates/template.xlsx` (contains marker cells)
- Output: `output/tasks.xlsx` (created/overwritten by the tool)

## MCP tool and prompt
- Tool: `write_task_list(main_tasks, support_tasks, template_path?, output_path?)`
  - Writes numbered tasks under the marker columns, preserving the row style beneath each marker.
- Prompt: `random_task_list` (registered in the server)
  - Asks the model to invent tasks and call `write_task_list`.
  - Ready-to-paste Copilot text lives in `prompts/copilot_random_task_prompt.txt`.

## Using with GitHub Copilot / MCP client
1. Start the MCP server (`python -m src.app`).
2. In your MCP-enabled client (e.g., Copilot Chat), select the server and choose the `random_task_list` prompt, or paste the text from `prompts/copilot_random_task_prompt.txt`.
3. The model will generate tasks and invoke `write_task_list`; check `output/tasks.xlsx` for results.

## Customizing
- Change template/output paths by passing `template_path`/`output_path` to `write_task_list` or by editing the prompt text.
- Task counts can be adjusted via the prompt arguments (`main_count`, `support_count`) or by editing the Copilot prompt file.

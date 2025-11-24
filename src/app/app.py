from __future__ import annotations

from pathlib import Path

from fastmcp import FastMCP

from ..helpers.excel import write_task_lists_to_excel


# BASE_DIR = Path(__file__).resolve().parents[2]
# DEFAULT_TEMPLATE = BASE_DIR / "templates" / "template.xlsx"
# DEFAULT_OUTPUT = BASE_DIR / "output" / "tasks.xlsx"

from . import settings


server = FastMCP(
    name="excel-task-server",
    instructions=(
        "Generate numbered main/support tasks and write them into the Excel template "
        "using marker cells (marker[main_task], marker[support_task]). "
        "Use the write_task_list tool to fill templates/template.xlsx and preserve formatting."
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
    template = template_path or str(DEFAULT_TEMPLATE)
    output = output_path or str(DEFAULT_OUTPUT)
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

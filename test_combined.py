"""Test script to create a complete combined Excel file."""

from src.helpers.combined_excel import write_combined_excel
from src.helpers.search import search_in_folder
from src.helpers.code import capture_multiple_blocks

# 1. Capture function (search_in_folder from search.py, lines 73-130)
blocks = capture_multiple_blocks('src/helpers/search.py', [(73, 130)])
print(f"Captured {len(blocks)} code blocks")

# 2. Search for "Monokai" keyword
summary = search_in_folder('src', 'Monokai', max_results=20)
print(f"Found {summary.total_matches} matches in {summary.files_with_matches} files")

# 3. Diff (simulating git commit changes)
old_code = """def _apply_code_style(ws, start_row: int, code_block: CodeBlock) -> int:
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", fill_type="solid")
    code_font = Font(name="Consolas", size=10)
    line_num_font = Font(name="Consolas", size=10, color="888888")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
    )"""

new_code = """def _apply_code_style(ws, start_row: int, code_block: CodeBlock, use_dark_theme: bool = True) -> int:
    # Theme colors - Monokai dark
    if use_dark_theme:
        bg_color = "272822"  # Monokai dark background
        header_bg = "1E1F1C"
        line_num_color = "75715E"
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color=header_bg, fill_type="solid")
    code_font = Font(name="Consolas", size=10, color=default_text_color)
    line_num_font = Font(name="Consolas", size=10, color=line_num_color)
    border = Border(
        left=Side(style="thin", color="444444"),
        right=Side(style="thin", color="444444"),
    )"""

# 4. PlantUML diagram
puml = """@startuml
title MCP Excel Server Architecture

actor User
participant "MCP Client" as Client
participant "search_in_folder" as Search  
participant "compute_diff" as Diff
participant "capture_function" as Capture
participant "Excel Writer" as Excel

User -> Client: Request combined doc
Client -> Capture: capture function
Client -> Diff: compare code versions
Client -> Search: search for keyword
Client -> Excel: write_combined_excel
Excel --> Client: all_in_one.xlsx
Client --> User: Done!
@enduml"""

# Create combined Excel with ALL sheets
path = write_combined_excel(
    'output/complete_demo.xlsx',
    code_blocks=blocks,
    diff_old=old_code,
    diff_new=new_code,
    diff_file_path='src/helpers/code_excel.py',
    search_summary=summary,
    puml_code=puml,
)

print(f"\n✅ Created: {path}")
print("📊 Sheets: Code, Diff, Search, Diagram")

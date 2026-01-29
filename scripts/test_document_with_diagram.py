"""Test script for documenting code.py with sequence diagram."""

from src.helpers.code_excel import write_code_and_diagram_to_excel
from src.helpers.code import capture_multiple_blocks

# Key functions to capture from src/helpers/code.py:
# 1. CodeBlock class: lines 9-16
# 2. capture_code_block: lines 55-105
# 3. format_code_block: lines 108-139
# 4. capture_multiple_blocks: lines 142-157

ranges = [
    (9, 16),    # CodeBlock class
    (55, 105),  # capture_code_block function
    (108, 139), # format_code_block function
    (142, 157), # capture_multiple_blocks function
]

blocks = capture_multiple_blocks('src/helpers/code.py', ranges)

puml_code = """@startuml
title Code Capture Flow

participant Client
participant "capture_code_block" as Capture
participant "Path" as Path
participant "File" as File
participant "detect_language" as DetectLang

== Capture Code Block ==
Client -> Capture: capture_code_block(file_path, start_line, end_line)
activate Capture

Capture -> Path: Path(file_path)
Path --> Capture: path object

Capture -> Path: path.exists()
Path --> Capture: True/False

alt File not found
    Capture --> Client: FileNotFoundError
end

Capture -> Capture: Validate line numbers

Capture -> File: open(path, "r")
activate File
File --> Capture: file handle

Capture -> File: readlines()
File --> Capture: lines[]
deactivate File

Capture -> Capture: Extract lines[start-1:end]

Capture -> DetectLang: detect_language(path)
activate DetectLang
DetectLang -> Path: path.suffix.lower()
Path --> DetectLang: extension
DetectLang --> Capture: language string
deactivate DetectLang

Capture --> Client: CodeBlock(file_path, start, end, code, language)
deactivate Capture

== Format Code Block ==
Client -> "format_code_block" as Format: format_code_block(code_block)
activate Format
Format -> Format: Split code into lines
Format -> Format: Add line numbers
Format -> Format: Build markdown
Format --> Client: Formatted string
deactivate Format
@enduml"""

result = write_code_and_diagram_to_excel(
    code_blocks=blocks,
    puml_code=puml_code,
    output_path='output/code_capture_documentation.xlsx',
    server_url='http://localhost:8080'
)

print(f"✅ Created: {result}")
print(f"\n📦 Captured {len(blocks)} code blocks:")
for i, b in enumerate(blocks, 1):
    print(f"   {i}. Lines {b.start_line}-{b.end_line}")

print(f"\n📊 Excel file has 2 sheets:")
print(f"   - Sheet 1: Code Blocks (with line numbers)")
print(f"   - Sheet 2: Sequence Diagram (PlantUML code + image)")

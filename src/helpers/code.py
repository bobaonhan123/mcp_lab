"""Code capture helper for extracting and documenting code blocks from files."""

from __future__ import annotations

from pathlib import Path
from typing import NamedTuple


class CodeBlock(NamedTuple):
    """Represents a captured code block."""
    file_path: str
    start_line: int
    end_line: int
    code: str
    language: str = ""


def detect_language(file_path: str | Path) -> str:
    """Detect programming language from file extension."""
    path = Path(file_path)
    ext = path.suffix.lower()
    
    language_map = {
        ".py": "python",
        ".js": "javascript",
        ".ts": "typescript",
        ".jsx": "javascript",
        ".tsx": "typescript",
        ".java": "java",
        ".cpp": "cpp",
        ".c": "c",
        ".h": "c",
        ".hpp": "cpp",
        ".cs": "csharp",
        ".rb": "ruby",
        ".go": "go",
        ".rs": "rust",
        ".php": "php",
        ".sql": "sql",
        ".html": "html",
        ".css": "css",
        ".json": "json",
        ".xml": "xml",
        ".yaml": "yaml",
        ".yml": "yaml",
        ".md": "markdown",
        ".sh": "bash",
        ".bat": "batch",
        ".ps1": "powershell",
    }
    
    return language_map.get(ext, "")


def capture_code_block(
    file_path: str | Path,
    start_line: int,
    end_line: int | None = None,
) -> CodeBlock:
    """Capture a code block from a file with line numbers.
    
    Args:
        file_path: Path to the source file
        start_line: Starting line number (1-indexed)
        end_line: Ending line number (1-indexed, inclusive). If None, captures single line.
        
    Returns:
        CodeBlock with file path, line numbers, and code content
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        ValueError: If line numbers are invalid
    """
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    
    if start_line < 1:
        raise ValueError(f"start_line must be >= 1, got {start_line}")
    
    if end_line is None:
        end_line = start_line
    elif end_line < start_line:
        raise ValueError(f"end_line ({end_line}) must be >= start_line ({start_line})")
    
    # Read file and capture lines
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    
    total_lines = len(lines)
    
    if start_line > total_lines:
        raise ValueError(f"start_line {start_line} exceeds file length {total_lines}")
    
    if end_line > total_lines:
        raise ValueError(f"end_line {end_line} exceeds file length {total_lines}")
    
    # Extract code block (convert 1-indexed to 0-indexed)
    code_lines = lines[start_line - 1 : end_line]
    code = "".join(code_lines)
    
    return CodeBlock(
        file_path=str(path),
        start_line=start_line,
        end_line=end_line,
        code=code,
        language=detect_language(path),
    )


def format_code_block(code_block: CodeBlock, include_line_numbers: bool = True) -> str:
    """Format a code block for display or documentation.
    
    Args:
        code_block: The code block to format
        include_line_numbers: Whether to include line numbers
        
    Returns:
        Formatted code block as string
    """
    lines = code_block.code.rstrip("\n").split("\n")
    
    # Build header
    header = f"File: {code_block.file_path}\n"
    header += f"Lines {code_block.start_line}-{code_block.end_line}\n"
    
    # Build code with optional line numbers
    if include_line_numbers:
        max_line_width = len(str(code_block.end_line))
        formatted_lines = []
        for i, line in enumerate(lines, start=code_block.start_line):
            line_num = str(i).rjust(max_line_width)
            formatted_lines.append(f"{line_num} | {line}")
        code_str = "\n".join(formatted_lines)
    else:
        code_str = code_block.code.rstrip("\n")
    
    # Build markdown code block
    markdown = f"```{code_block.language}\n{code_str}\n```"
    
    return f"{header}\n{markdown}"


def capture_multiple_blocks(
    file_path: str | Path,
    ranges: list[tuple[int, int | None]],
) -> list[CodeBlock]:
    """Capture multiple code blocks from the same file.
    
    Args:
        file_path: Path to the source file
        ranges: List of (start_line, end_line) tuples
        
    Returns:
        List of CodeBlock objects
    """
    blocks = []
    for start, end in ranges:
        block = capture_code_block(file_path, start, end)
        blocks.append(block)
    return blocks


def create_code_documentation(
    blocks: list[CodeBlock] | CodeBlock,
    title: str = "Code Documentation",
) -> str:
    """Create formatted documentation from code blocks.
    
    Args:
        blocks: Single CodeBlock or list of CodeBlocks
        title: Document title
        
    Returns:
        Formatted documentation as markdown string
    """
    if isinstance(blocks, CodeBlock):
        blocks = [blocks]
    
    doc = f"# {title}\n\n"
    
    for i, block in enumerate(blocks, start=1):
        doc += f"## Block {i}\n"
        doc += format_code_block(block, include_line_numbers=True)
        doc += "\n\n"
    
    return doc

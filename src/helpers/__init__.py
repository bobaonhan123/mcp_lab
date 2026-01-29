"""Helper modules for Excel, PlantUML, and Code operations."""

from .excel import write_task_lists_to_excel, MarkerNotFoundError
from .plantuml import (
    generate_plantuml_image,
    write_plantuml_to_excel,
    write_plantuml_image_only,
)
from .code import (
    capture_code_block,
    capture_multiple_blocks,
    format_code_block,
    create_code_documentation,
    CodeBlock,
)
from .code_excel import (
    write_code_blocks_to_excel,
    write_code_and_diagram_to_excel,
    capture_and_write_to_excel,
)
from .code_ast import (
    analyze_python_file,
    capture_function,
    capture_class,
    capture_method,
    capture_all_functions,
    capture_all_classes,
    capture_by_names,
    get_file_summary,
    CodeElement,
)
from .diff import (
    compute_diff,
    write_diff_to_excel,
    format_diff_text,
    DiffBlock,
    DiffLine,
)
from .search import (
    search_in_folder,
    write_search_to_excel,
    format_search_results,
    SearchMatch,
    SearchResult,
    SearchSummary,
)
from .combined_excel import write_combined_excel

__all__ = [
    "write_task_lists_to_excel",
    "MarkerNotFoundError",
    "generate_plantuml_image",
    "write_plantuml_to_excel",
    "write_plantuml_image_only",
    "capture_code_block",
    "capture_multiple_blocks",
    "format_code_block",
    "create_code_documentation",
    "CodeBlock",
    "write_code_blocks_to_excel",
    "write_code_and_diagram_to_excel",
    "capture_and_write_to_excel",
    "analyze_python_file",
    "capture_function",
    "capture_class",
    "capture_method",
    "capture_all_functions",
    "capture_all_classes",
    "capture_by_names",
    "get_file_summary",
    "CodeElement",
    # Diff
    "compute_diff",
    "write_diff_to_excel",
    "format_diff_text",
    "DiffBlock",
    "DiffLine",
    # Search
    "search_in_folder",
    "write_search_to_excel",
    "format_search_results",
    "SearchMatch",
    "SearchResult",
    "SearchSummary",
]


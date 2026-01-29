"""AST-based code analysis for automatic function/class detection."""

from __future__ import annotations

import ast
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from .code import CodeBlock, capture_code_block, detect_language


@dataclass
class CodeElement:
    """Represents a code element (function, class, method) with metadata."""
    name: str
    element_type: Literal["function", "class", "method", "async_function", "async_method"]
    start_line: int
    end_line: int
    docstring: str | None = None
    decorators: list[str] | None = None
    parent_class: str | None = None


def analyze_python_file(file_path: str | Path) -> list[CodeElement]:
    """Analyze a Python file and extract all functions and classes with their line ranges.
    
    Args:
        file_path: Path to the Python file
        
    Returns:
        List of CodeElement objects with name, type, and line ranges
    """
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    
    with open(path, "r", encoding="utf-8") as f:
        source = f.read()
    
    try:
        tree = ast.parse(source)
    except SyntaxError as e:
        raise ValueError(f"Syntax error in {file_path}: {e}")
    
    elements = []
    
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            # Extract class
            decorators = [_get_decorator_name(d) for d in node.decorator_list]
            docstring = ast.get_docstring(node)
            
            elements.append(CodeElement(
                name=node.name,
                element_type="class",
                start_line=node.lineno,
                end_line=node.end_lineno or node.lineno,
                docstring=docstring,
                decorators=decorators if decorators else None,
            ))
            
            # Extract methods within the class
            for item in node.body:
                if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    method_decorators = [_get_decorator_name(d) for d in item.decorator_list]
                    method_docstring = ast.get_docstring(item)
                    
                    element_type = "async_method" if isinstance(item, ast.AsyncFunctionDef) else "method"
                    
                    elements.append(CodeElement(
                        name=item.name,
                        element_type=element_type,
                        start_line=item.lineno,
                        end_line=item.end_lineno or item.lineno,
                        docstring=method_docstring,
                        decorators=method_decorators if method_decorators else None,
                        parent_class=node.name,
                    ))
        
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            # Skip methods (already handled above)
            if _is_method(node, tree):
                continue
            
            decorators = [_get_decorator_name(d) for d in node.decorator_list]
            docstring = ast.get_docstring(node)
            
            element_type = "async_function" if isinstance(node, ast.AsyncFunctionDef) else "function"
            
            elements.append(CodeElement(
                name=node.name,
                element_type=element_type,
                start_line=node.lineno,
                end_line=node.end_lineno or node.lineno,
                docstring=docstring,
                decorators=decorators if decorators else None,
            ))
    
    # Sort by start line
    elements.sort(key=lambda e: e.start_line)
    
    return elements


def _get_decorator_name(decorator: ast.expr) -> str:
    """Extract decorator name from AST node."""
    if isinstance(decorator, ast.Name):
        return decorator.id
    elif isinstance(decorator, ast.Attribute):
        return f"{_get_decorator_name(decorator.value)}.{decorator.attr}"
    elif isinstance(decorator, ast.Call):
        return _get_decorator_name(decorator.func)
    return str(decorator)


def _is_method(node: ast.FunctionDef | ast.AsyncFunctionDef, tree: ast.Module) -> bool:
    """Check if a function node is a method inside a class."""
    for parent in ast.walk(tree):
        if isinstance(parent, ast.ClassDef):
            for item in parent.body:
                if item is node:
                    return True
    return False


def capture_function(
    file_path: str | Path,
    function_name: str,
    include_decorators: bool = True,
) -> CodeBlock:
    """Capture a function by name, automatically finding its line range.
    
    Args:
        file_path: Path to the Python file
        function_name: Name of the function to capture
        include_decorators: Whether to include decorator lines
        
    Returns:
        CodeBlock with the function code
    """
    elements = analyze_python_file(file_path)
    
    for elem in elements:
        if elem.name == function_name and elem.element_type in ("function", "async_function"):
            return capture_code_block(file_path, elem.start_line, elem.end_line)
    
    raise ValueError(f"Function '{function_name}' not found in {file_path}")


def capture_class(
    file_path: str | Path,
    class_name: str,
) -> CodeBlock:
    """Capture a class by name, automatically finding its line range.
    
    Args:
        file_path: Path to the Python file
        class_name: Name of the class to capture
        
    Returns:
        CodeBlock with the class code
    """
    elements = analyze_python_file(file_path)
    
    for elem in elements:
        if elem.name == class_name and elem.element_type == "class":
            return capture_code_block(file_path, elem.start_line, elem.end_line)
    
    raise ValueError(f"Class '{class_name}' not found in {file_path}")


def capture_method(
    file_path: str | Path,
    class_name: str,
    method_name: str,
) -> CodeBlock:
    """Capture a method by class and method name.
    
    Args:
        file_path: Path to the Python file
        class_name: Name of the class containing the method
        method_name: Name of the method to capture
        
    Returns:
        CodeBlock with the method code
    """
    elements = analyze_python_file(file_path)
    
    for elem in elements:
        if (elem.name == method_name and 
            elem.element_type in ("method", "async_method") and
            elem.parent_class == class_name):
            return capture_code_block(file_path, elem.start_line, elem.end_line)
    
    raise ValueError(f"Method '{class_name}.{method_name}' not found in {file_path}")


def capture_all_functions(
    file_path: str | Path,
    include_methods: bool = False,
) -> list[CodeBlock]:
    """Capture all functions (and optionally methods) from a file.
    
    Args:
        file_path: Path to the Python file
        include_methods: Whether to include class methods
        
    Returns:
        List of CodeBlock objects for all functions
    """
    elements = analyze_python_file(file_path)
    blocks = []
    
    for elem in elements:
        if elem.element_type in ("function", "async_function"):
            blocks.append(capture_code_block(file_path, elem.start_line, elem.end_line))
        elif include_methods and elem.element_type in ("method", "async_method"):
            blocks.append(capture_code_block(file_path, elem.start_line, elem.end_line))
    
    return blocks


def capture_all_classes(file_path: str | Path) -> list[CodeBlock]:
    """Capture all classes from a file.
    
    Args:
        file_path: Path to the Python file
        
    Returns:
        List of CodeBlock objects for all classes
    """
    elements = analyze_python_file(file_path)
    blocks = []
    
    for elem in elements:
        if elem.element_type == "class":
            blocks.append(capture_code_block(file_path, elem.start_line, elem.end_line))
    
    return blocks


def capture_by_names(
    file_path: str | Path,
    names: list[str],
) -> list[CodeBlock]:
    """Capture multiple functions/classes by their names.
    
    Args:
        file_path: Path to the Python file
        names: List of function or class names to capture
        
    Returns:
        List of CodeBlock objects
    """
    elements = analyze_python_file(file_path)
    blocks = []
    
    for name in names:
        for elem in elements:
            if elem.name == name:
                blocks.append(capture_code_block(file_path, elem.start_line, elem.end_line))
                break
        else:
            raise ValueError(f"'{name}' not found in {file_path}")
    
    return blocks


def get_file_summary(file_path: str | Path) -> str:
    """Get a summary of all code elements in a file.
    
    Args:
        file_path: Path to the Python file
        
    Returns:
        Formatted summary string
    """
    elements = analyze_python_file(file_path)
    
    lines = [f"# Code Summary: {file_path}\n"]
    
    current_class = None
    
    for elem in elements:
        if elem.element_type == "class":
            current_class = elem.name
            lines.append(f"\n## Class: {elem.name} (lines {elem.start_line}-{elem.end_line})")
            if elem.docstring:
                lines.append(f"   {elem.docstring.split(chr(10))[0]}")
        
        elif elem.element_type in ("method", "async_method"):
            prefix = "async " if elem.element_type == "async_method" else ""
            lines.append(f"   - {prefix}{elem.name}() (lines {elem.start_line}-{elem.end_line})")
            if elem.docstring:
                lines.append(f"     {elem.docstring.split(chr(10))[0]}")
        
        elif elem.element_type in ("function", "async_function"):
            prefix = "async " if elem.element_type == "async_function" else ""
            lines.append(f"\n### {prefix}Function: {elem.name}() (lines {elem.start_line}-{elem.end_line})")
            if elem.docstring:
                lines.append(f"   {elem.docstring.split(chr(10))[0]}")
    
    return "\n".join(lines)

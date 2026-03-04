"""MCP DocGen Server - Document generation and extraction tools for AI agents."""

import sys
import os

# Windows defaults sys.stdio to cp1252, corrompendo caracteres não-ASCII
# no transporte MCP via stdio. Forçar UTF-8 antes de qualquer import.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")
if hasattr(sys.stdin, "reconfigure"):
    sys.stdin.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastmcp import FastMCP

from handlers.docx_handler import create_docx, read_docx
from handlers.pdf_handler import create_pdf, read_pdf
from handlers.excel_handler import create_excel, read_excel
from handlers.markdown_handler import markdown_to_document
from handlers.chart_handler import create_chart, create_chart_document
from handlers.txt_handler import create_txt, read_txt, append_txt
from handlers.zip_handler import create_zip, read_zip
from utils.file_utils import list_files

mcp = FastMCP(
    name="DocGen",
    instructions=(
        "MCP server for generating and reading documents (Word, PDF, Excel, TXT) and creating ZIP archives. "
        "All create tools require an output_path parameter with the absolute file path where the document will be saved. "
        "Files at existing paths will be overwritten. "
        "Use create_docx/create_pdf/create_excel to generate documents from structured sections. "
        "Use read_docx/read_pdf/read_excel to extract content from existing files. "
        "Use markdown_to_document to convert raw markdown into docx or pdf. "
        "Use create_chart to generate charts (bar, line, pie, scatter, area, horizontal_bar) as PNG images. "
        "Use create_chart_document to generate a chart and embed it directly in a docx or pdf. "
        "Use create_txt to write plain text files, read_txt to read them, append_txt to append content. "
        "Use create_zip to bundle files into a .zip archive, read_zip to inspect archive contents. "
        "Use list_files to list files in any directory. "
        "Use create_folder to create a directory (and all parents) at a given path. "
        "Sections follow a unified schema: heading, paragraph, table, image, code_block, list, page_break. "
        "Images are passed as absolute file paths and are embedded into the documents. "
        "Write operations are blocked on sensitive system directories (Windows, Program Files, .ssh, .aws, etc.)."
    ),
)


# ─── DOCX ────────────────────────────────────────────────────────────────────


@mcp.tool(
    name="create_docx",
    description=(
        "Generate a Word (.docx) document from structured sections. "
        "Sections: heading, paragraph, table, image, code_block, list, page_break. "
        "Images are embedded from file paths. Returns the output file path."
    ),
    tags={"document", "word", "docx", "generate"},
)
def tool_create_docx(
    output_path: str,
    title: str,
    sections: list[dict],
) -> str:
    """Create a Word document.

    Args:
        output_path: Absolute path for the output file (e.g. "C:/docs/report.docx"). Overwrites if exists.
        title: Document title (shown as main heading).
        sections: List of section dicts. Each has "type" and type-specific fields.
            - heading: {"type":"heading", "text":"...", "level":1}
            - paragraph: {"type":"paragraph", "text":"...", "bold":false, "italic":false}
            - table: {"type":"table", "headers":["A","B"], "rows":[["1","2"]], "style":"Light Grid Accent 1"}
            - image: {"type":"image", "path":"C:/img.png", "width_inches":5.0, "caption":"..."}
            - code_block: {"type":"code_block", "code":"print('hi')", "language":"python"}
            - list: {"type":"list", "items":["a","b"], "ordered":false}
            - page_break: {"type":"page_break"}
    """
    path = create_docx(title, sections, output_path)
    return f"Documento Word gerado: {path}"


@mcp.tool(
    name="read_docx",
    description=(
        "Extract text, tables, images count and metadata from a Word (.docx) file."
    ),
    tags={"document", "word", "docx", "read", "extract"},
)
def tool_read_docx(file_path: str) -> dict:
    """Read and extract content from a Word document.

    Args:
        file_path: Absolute path to the .docx file.
    """
    return read_docx(file_path)


# ─── PDF ─────────────────────────────────────────────────────────────────────


@mcp.tool(
    name="create_pdf",
    description=(
        "Generate a PDF document from structured sections. "
        "Same section schema as create_docx. Returns the output file path."
    ),
    tags={"document", "pdf", "generate"},
)
def tool_create_pdf(
    output_path: str,
    title: str,
    sections: list[dict],
) -> str:
    """Create a PDF document.

    Args:
        output_path: Absolute path for the output file (e.g. "C:/docs/report.pdf"). Overwrites if exists.
        title: Document title.
        sections: List of section dicts (same schema as create_docx).
    """
    path = create_pdf(title, sections, output_path)
    return f"Documento PDF gerado: {path}"


# ─── EXCEL ───────────────────────────────────────────────────────────────────


@mcp.tool(
    name="create_excel",
    description=(
        "Generate an Excel (.xlsx) workbook with multiple sheets, headers, rows, "
        "auto-sized columns, and optional embedded images."
    ),
    tags={"document", "excel", "xlsx", "generate"},
)
def tool_create_excel(
    output_path: str,
    title: str,
    sheets: list[dict],
) -> str:
    """Create an Excel workbook.

    Args:
        output_path: Absolute path for the output file (e.g. "C:/docs/data.xlsx"). Overwrites if exists.
        title: Document title.
        sheets: List of sheet definitions:
            - name: Sheet tab name
            - headers: Column header list
            - rows: List of row lists
            - column_types: Optional list of column type hints, one per column.
                Supported values: "string", "number", "date", "auto" (default).
                When "date", string values like "15/08/2028" or "2028-08-15" are
                parsed into native Excel date cells.
                When "number", values are forced to numeric type.
                When "string", values are kept as text.
                When "auto", heuristic conversion is applied (numbers detected automatically).
            - date_format: Optional display format for date columns (default "DD/MM/YYYY").
                Supported: "DD/MM/YYYY", "YYYY-MM-DD", "DD-MM-YYYY", "MM/DD/YYYY", "DD.MM.YYYY".
            - column_widths: Optional list of column widths
            - images: Optional list of {"path":"...", "cell":"A1"}
    """
    path = create_excel(title, sheets, output_path)
    return f"Planilha Excel gerada: {path}"


@mcp.tool(
    name="read_excel",
    description="Extract headers, rows, and sheet info from an Excel (.xlsx) file.",
    tags={"document", "excel", "xlsx", "read", "extract"},
)
def tool_read_excel(file_path: str, sheet_name: str | None = None) -> dict:
    """Read and extract content from an Excel workbook.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Optional specific sheet to read. Reads all if omitted.
    """
    return read_excel(file_path, sheet_name)


@mcp.tool(
    name="read_pdf",
    description="Extract text, tables, and metadata from a PDF file.",
    tags={"document", "pdf", "read", "extract"},
)
def tool_read_pdf(file_path: str, pages: list[int] | None = None) -> dict:
    """Read and extract content from a PDF.

    Args:
        file_path: Absolute path to the .pdf file.
        pages: Optional list of 0-based page indices. Reads all if omitted.
    """
    return read_pdf(file_path, pages)


# ─── MARKDOWN ────────────────────────────────────────────────────────────────


@mcp.tool(
    name="markdown_to_document",
    description=(
        "Convert raw Markdown text into a Word (.docx) or PDF document. "
        "Supports headings, bold/italic, tables, code blocks, lists, images (![caption](path)), "
        "and page breaks (---). Returns the output file path."
    ),
    tags={"document", "markdown", "convert", "generate"},
)
def tool_markdown_to_document(
    output_path: str,
    markdown: str,
    output_format: str = "docx",
    title: str | None = None,
) -> str:
    """Convert Markdown to a Word or PDF document.

    Args:
        output_path: Absolute path for the output file. Overwrites if exists.
        markdown: Raw markdown text to convert.
        output_format: "docx" or "pdf". Defaults to "docx".
        title: Optional document title. Auto-detected from first heading if omitted.
    """
    path = markdown_to_document(markdown, output_format, title, output_path)
    fmt = "PDF" if output_format.lower() == "pdf" else "Word"
    return f"Documento {fmt} gerado a partir de Markdown: {path}"


# ─── CHARTS ──────────────────────────────────────────────────────────────────


@mcp.tool(
    name="create_chart",
    description=(
        "Generate a chart as a PNG image using matplotlib. "
        "Supported types: bar, horizontal_bar, line, pie, scatter, area. "
        "Returns the path to the generated PNG."
    ),
    tags={"chart", "graph", "matplotlib", "generate"},
)
def tool_create_chart(
    output_path: str,
    chart_type: str,
    data: dict,
    title: str = "",
    xlabel: str = "",
    ylabel: str = "",
    width: float = 8.0,
    height: float = 5.0,
) -> str:
    """Generate a chart image.

    Args:
        output_path: Absolute path for the PNG output. Overwrites if exists.
        chart_type: One of "bar", "horizontal_bar", "line", "pie", "scatter", "area".
        data: Chart data structure.
            For bar/line/area: {"labels": ["A","B"], "datasets": [{"label": "S1", "values": [10,20]}]}
            For pie: {"labels": ["A","B"], "values": [30,70]}
            For scatter: {"datasets": [{"label": "S1", "x": [1,2], "y": [3,4]}]}
        title: Chart title displayed above the chart.
        xlabel: X-axis label.
        ylabel: Y-axis label.
        width: Figure width in inches (default 8).
        height: Figure height in inches (default 5).
    """
    path = create_chart(chart_type, data, title, xlabel, ylabel, width, height, output_path)
    return f"Gráfico '{chart_type}' gerado: {path}"


@mcp.tool(
    name="create_chart_document",
    description=(
        "Generate a chart and embed it directly in a Word (.docx) or PDF document. "
        "Optionally add extra sections (text, tables) after the chart."
    ),
    tags={"chart", "document", "generate"},
)
def tool_create_chart_document(
    output_path: str,
    chart_type: str,
    data: dict,
    title: str = "",
    xlabel: str = "",
    ylabel: str = "",
    output_format: str = "docx",
    extra_sections: list[dict] | None = None,
) -> str:
    """Generate a chart embedded in a document.

    Args:
        output_path: Absolute path for the output document. Overwrites if exists.
        chart_type: One of "bar", "horizontal_bar", "line", "pie", "scatter", "area".
        data: Chart data (same schema as create_chart).
        title: Chart and document title.
        xlabel: X-axis label.
        ylabel: Y-axis label.
        output_format: "docx" or "pdf". Defaults to "docx".
        extra_sections: Optional list of sections to add after the chart.
    """
    path = create_chart_document(
        chart_type, data, title, xlabel, ylabel,
        output_format, extra_sections, output_path
    )
    fmt = "PDF" if output_format.lower() == "pdf" else "Word"
    return f"Documento {fmt} com gráfico '{chart_type}' gerado: {path}"


# ─── TXT ──────────────────────────────────────────────────────────────────────


@mcp.tool(
    name="create_txt",
    description=(
        "Create a plain text (.txt) file with the given content. "
        "Returns the output file path."
    ),
    tags={"document", "txt", "text", "generate", "write"},
)
def tool_create_txt(
    output_path: str,
    content: str,
    encoding: str = "utf-8",
) -> str:
    """Create a plain text file.

    Args:
        output_path: Absolute path for the output file (e.g. "C:/docs/notes.txt"). Overwrites if exists.
        content: Text content to write into the file.
        encoding: File encoding. Defaults to "utf-8".
    """
    path = create_txt(content, output_path, encoding)
    return f"Arquivo TXT criado: {path}"


@mcp.tool(
    name="read_txt",
    description="Read a plain text file and return its content, line count, and size.",
    tags={"document", "txt", "text", "read", "extract"},
)
def tool_read_txt(file_path: str, encoding: str = "utf-8") -> dict:
    """Read a plain text file.

    Args:
        file_path: Absolute path to the .txt file.
        encoding: File encoding. Defaults to "utf-8".
    """
    return read_txt(file_path, encoding)


@mcp.tool(
    name="append_txt",
    description=(
        "Append content to an existing text file. Creates the file if it doesn't exist. "
        "Returns the file path."
    ),
    tags={"document", "txt", "text", "write", "append"},
)
def tool_append_txt(
    file_path: str,
    content: str,
    encoding: str = "utf-8",
) -> str:
    """Append text to a file.

    Args:
        file_path: Absolute path to the .txt file.
        content: Text content to append.
        encoding: File encoding. Defaults to "utf-8".
    """
    path = append_txt(file_path, content, encoding)
    return f"Conteúdo adicionado ao arquivo: {path}"


# ─── ZIP ──────────────────────────────────────────────────────────────────────


@mcp.tool(
    name="create_zip",
    description=(
        "Create a ZIP archive (.zip) from a list of files or directories. "
        "Directories are added recursively. Returns the output .zip path."
    ),
    tags={"archive", "zip", "compress", "generate"},
)
def tool_create_zip(
    output_path: str,
    file_paths: list[str],
    compression: str = "deflated",
) -> str:
    """Create a ZIP archive.

    Args:
        output_path: Absolute path for the .zip file (e.g. "C:/docs/files.zip"). Overwrites if exists.
        file_paths: List of absolute paths to files or directories to include.
            Directories are added recursively with relative paths preserved.
        compression: "deflated" (default, smaller size) or "stored" (no compression, faster).
    """
    path = create_zip(file_paths, output_path, compression)
    return f"Arquivo ZIP criado: {path}"


@mcp.tool(
    name="read_zip",
    description="List contents of a ZIP archive without extracting. Shows file names, sizes, and compression info.",
    tags={"archive", "zip", "read", "inspect"},
)
def tool_read_zip(file_path: str) -> dict:
    """Inspect a ZIP archive.

    Args:
        file_path: Absolute path to the .zip file.
    """
    return read_zip(file_path)


# ─── UTILITY ─────────────────────────────────────────────────────────────────


@mcp.tool(
    name="list_files",
    description="List files in a directory with name, path, size and modification date.",
    tags={"utility", "files"},
)
def tool_list_files(directory: str, extension: str | None = None) -> list[dict]:
    """List files in a directory.

    Args:
        directory: Absolute path to the directory to list.
        extension: Optional filter by extension (e.g. "docx", "pdf", "xlsx", "txt", "zip").
    """
    return list_files(directory, extension)


@mcp.tool(
    name="create_folder",
    description=(
        "Create a directory (folder) at the given path. "
        "Creates all intermediate parent directories if needed. "
        "Does nothing if the folder already exists. "
        "Write operations are blocked on sensitive system directories. "
        "Returns the absolute path of the created folder."
    ),
    tags={"utility", "files", "folder", "directory"},
)
def tool_create_folder(folder_path: str) -> str:
    """Create a directory at the given path.

    Args:
        folder_path: Absolute path of the folder to create (e.g. "C:/docs/my-folder").
            All intermediate parents are created automatically.
    """
    from utils.security import validate_write_path
    resolved = validate_write_path(folder_path)
    resolved.mkdir(parents=True, exist_ok=True)
    return f"Pasta criada: {resolved}"


# ─── ENTRY POINT ─────────────────────────────────────────────────────────────


if __name__ == "__main__":
    mcp.run(transport="stdio")

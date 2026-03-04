# DocGen MCP Server

A powerful **Model Context Protocol (MCP)** server that gives AI agents the ability to generate and read documents in multiple formats. Create professional Word documents, PDFs, Excel workbooks, PowerPoint presentations, charts, text files, and ZIP archives—all through a single, unified toolset.

## Why DocGen?

DocGen turns your AI assistant into a full-featured document engine. It exposes a rich set of tools so agents can produce **polished, structured output** instead of plain text: reports in DOCX or PDF, spreadsheets with multiple sheets, presentations from HTML, and charts as images or embedded in documents. All operations are **path-validated** and **security-conscious**, blocking writes to sensitive system directories.

---

## Features & Capabilities

### Document generation

- **Word (DOCX)** — Generate documents from structured sections: headings, paragraphs, **bold**/ *italic* text, tables, images, code blocks, and bulleted or numbered lists. Perfect for reports, specs, and formal documents.
- **PDF** — Same section schema as DOCX: titles, body text, tables, images, code, lists, and page breaks. Ideal for sharing and printing.
- **Excel (XLSX)** — Build workbooks with multiple sheets, custom headers, typed columns (string, number, date, auto), optional column widths, and embedded images. Supports common date formats.

### Presentations from HTML

- **PowerPoint (PPTX)** — Generate slides directly from **HTML**. Use `<section>` or `<div class="slide">` for multiple slides; support for **rich text** (bold, italic), **colors** via `<span style="color: #hex">`, **tables**, **lists**, **images** (file path or data URI), and **clickable links** (`<a href>`). Great for colorful, interactive decks without leaving your workflow.

### Markdown & charts

- **Markdown to document** — Convert raw Markdown into DOCX or PDF. Handles headings, bold/italic, tables, code blocks, lists, images, and page breaks (`---`).
- **Charts** — Create bar, horizontal bar, line, pie, scatter, and area charts as PNG, or embed them directly in a DOCX or PDF with optional extra sections.

### Files & utilities

- **Text files** — Create, read, and append plain text with configurable encoding.
- **ZIP archives** — Bundle files and directories (recursive) into a ZIP; list contents without extracting.
- **Files & folders** — List directory contents (optional filter by extension) and create directories (including parents).

### Security & reliability

- **Path validation** — All write operations validate paths and block traversal and symlink tricks.
- **Protected directories** — Writes to system and sensitive folders (e.g. Windows, Program Files, `.ssh`, `.aws`) are blocked.
- **Safe image handling** — Image paths are validated and restricted to supported formats.

---

## Requirements

- **Python 3.10+**
- Dependencies listed in `requirements.txt`

---

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd docgen-mcp-server

# Create a virtual environment (recommended)
python -m venv venv
# Windows (PowerShell):
.\venv\Scripts\Activate.ps1
# Linux/macOS:
# source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

---

## Using the MCP Server

### 1. Configure in Cursor or your MCP client

Add the server to your MCP configuration. Example for **Cursor** (in MCP settings or project config):

```json
{
  "mcpServers": {
    "docgen": {
      "command": "python",
      "args": ["-m", "server"],
      "cwd": "C:/path/to/docgen-mcp-server",
      "env": {}
    }
  }
}
```

Replace `C:/path/to/docgen-mcp-server` with the absolute path to the project folder.

**Using the venv Python executable:**

```json
{
  "mcpServers": {
    "docgen": {
      "command": "C:/path/to/docgen-mcp-server/venv/Scripts/python.exe",
      "args": ["-m", "server"],
      "cwd": "C:/path/to/docgen-mcp-server"
    }
  }
}
```

### 2. Run manually (stdio)

To run the server directly (e.g. for debugging):

```bash
# From the project root with venv activated
python -m server
```

The server uses **stdio** transport: it reads JSON-RPC from stdin and writes to stdout. Your MCP client starts this process and connects the pipes.

### 3. Available tools

Once connected, the client can call tools such as:

| Tool | Description |
|------|-------------|
| `create_docx` | Generate Word from sections (title, paragraphs, tables, images, code, lists). |
| `read_docx` | Extract text, tables, and metadata from a .docx. |
| `create_pdf` | Generate PDF with the same section schema as DOCX. |
| `read_pdf` | Extract text and tables from a PDF (optional page range). |
| `create_excel` | Generate an .xlsx workbook with multiple sheets, headers, rows, and optional images. |
| `read_excel` | Extract sheet names, headers, and rows from an .xlsx. |
| `create_pptx_from_html` | Generate PowerPoint from HTML: multiple slides, rich text, colors, tables, lists, images, links. |
| `markdown_to_document` | Convert Markdown to .docx or .pdf. |
| `create_chart` | Generate a chart (bar, line, pie, scatter, area, horizontal_bar) as PNG. |
| `create_chart_document` | Generate a chart and embed it in a docx or pdf, with optional extra sections. |
| `create_txt` / `read_txt` / `append_txt` | Create, read, and append plain text files. |
| `create_zip` / `read_zip` | Create ZIP archives and list their contents. |
| `list_files` | List files in a directory (optional extension filter). |
| `create_folder` | Create a directory and all parent directories. |

All creation tools require an `output_path` with the absolute path for the output file.

---

## Project structure

```
docgen-mcp-server/
├── server.py          # FastMCP entry point and tool registration
├── __main__.py        # Enables: python -m server
├── requirements.txt
├── handlers/          # Generation and reading logic by format
│   ├── docx_handler.py
│   ├── pdf_handler.py
│   ├── excel_handler.py
│   ├── pptx_handler.py
│   ├── markdown_handler.py
│   ├── chart_handler.py
│   ├── txt_handler.py
│   └── zip_handler.py
└── utils/
    ├── file_utils.py   # Output paths and directory listing
    └── security.py     # Path validation and blocked directories
```

---

## License

As defined in the project repository.

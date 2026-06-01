from mcp.server.fastmcp import FastMCP
from typing import Optional, List, Dict, Any
import json

from services import (
    read_doc_service, read_sheet_service, read_archive_service,
    write_doc_service, write_sheet_service, render_slide_service,
    render_page_service, patch_doc_service, patch_sheet_service,
    scan_dir_service, diff_file_service, bundle_zip_service
)


mcp = FastMCP("Docgen")

@mcp.tool(name="read_doc")
def read_doc(
    path: str,
    includeComments: Optional[bool] = None,
    extractImages: Optional[bool] = None,
    previewOnly: Optional[bool] = None,
    maxChars: Optional[int] = None
) -> str:
    opts = {
        "includeComments": includeComments,
        "extractImages": extractImages,
        "previewOnly": previewOnly,
        "maxChars": maxChars
    }
    res = read_doc_service(path, opts)
    text = res["markdown"]
    if res.get("note"):
        text += f"\n\n_{res['note']}_"
    return text

@mcp.tool(name="read_sheet")
def read_sheet(
    path: str,
    sheetName: Optional[str] = None,
    range: Optional[str] = None,
    asJson: bool = False,
    previewOnly: Optional[bool] = None,
    maxRows: Optional[int] = None
) -> str:
    opts = {
        "sheetName": sheetName,
        "range": range,
        "asJson": asJson,
        "previewOnly": previewOnly,
        "maxRows": maxRows
    }
    res = read_sheet_service(path, opts)
    if res["asJson"]:
        return json.dumps(res.get("jsonRows", []), indent=2)
    return res.get("markdown", "")

@mcp.tool(name="read_archive")
def read_archive(
    path: str,
    pattern: Optional[str] = None
) -> str:
    return read_archive_service(path, pattern)

@mcp.tool(name="write_doc")
def write_doc(
    path: str,
    type: str,
    content: str,
    contentFormat: Optional[str] = "markdown",
    templatePath: Optional[str] = None,
    mergeFields: Optional[Dict[str, Any]] = None
) -> str:
    opts = {
        "path": path,
        "type": type,
        "content": content,
        "contentFormat": contentFormat,
        "templatePath": templatePath,
        "mergeFields": mergeFields
    }
    return write_doc_service(opts)

@mcp.tool(name="write_sheet")
def write_sheet(
    path: str,
    data: List[Any],
    columns: Optional[Dict[str, str]] = None,
    freezePanes: Optional[bool] = None,
    append: Optional[bool] = None
) -> str:
    opts = {
        "path": path,
        "data": data,
        "columns": columns,
        "freezePanes": freezePanes,
        "append": append
    }
    return write_sheet_service(opts)

@mcp.tool(name="render_slide")
def render_slide(
    html: str,
    css: str,
    format: str,
    aspectRatio: str,
    outputPath: str
) -> str:
    opts = {
        "html": html,
        "css": css,
        "format": format,
        "aspectRatio": aspectRatio,
        "outputPath": outputPath
    }
    return render_slide_service(opts)

@mcp.tool(name="render_page")
def render_page(
    html: str,
    css: str,
    outputPath: str,
    generateTOC: Optional[bool] = None,
    margins: Optional[Dict[str, str]] = None
) -> str:
    opts = {
        "html": html,
        "css": css,
        "outputPath": outputPath,
        "generateTOC": generateTOC,
        "margins": margins
    }
    return render_page_service(opts)

@mcp.tool(name="patch_doc")
def patch_doc(
    path: str,
    action: str,
    payload: Any
) -> str:
    opts = {
        "path": path,
        "action": action,
        "payload": payload
    }
    return patch_doc_service(opts)

@mcp.tool(name="patch_sheet")
def patch_sheet(
    path: str,
    updates: List[Dict[str, Any]]
) -> str:
    return patch_sheet_service(path, updates)

@mcp.tool(name="scan_dir")
def scan_dir(
    path: str,
    query: str,
    recursive: bool,
    maxMatches: Optional[int] = None
) -> str:
    res = scan_dir_service(path, query, recursive, maxMatches)
    lines = res["matches"]
    text = "\n".join(lines)
    if res.get("truncated"):
        cap = maxMatches or 500
        text += f"\n\n_[Lista truncada: maximo {cap} correspondencias]_"
    return text if text else "Nenhuma correspondencia."

@mcp.tool(name="diff_file")
def diff_file(
    pathA: str,
    pathB: str,
    mode: str
) -> str:
    return diff_file_service(pathA, pathB, mode)

@mcp.tool(name="bundle_zip")
def bundle_zip(
    files: List[str],
    outputName: str
) -> str:
    return bundle_zip_service(files, outputName)

def main():
    mcp.run()

if __name__ == "__main__":
    main()

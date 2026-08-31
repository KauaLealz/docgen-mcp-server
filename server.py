from mcp.server.fastmcp import FastMCP
from typing import Optional, List, Dict, Any
import json

from services import (
    read_doc_service, read_sheet_service, read_archive_service,
    write_doc_service, write_sheet_service, render_slide_service,
    render_page_service, render_image_service, patch_doc_service,
    patch_sheet_service, scan_dir_service, diff_file_service,
    bundle_zip_service
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
    mergeFields: Optional[Dict[str, Any]] = None,
    css: Optional[str] = None,
    headerHtml: Optional[str] = None,
    footerHtml: Optional[str] = None,
    margins: Optional[Dict[str, str]] = None,
    fonts: Optional[List[Dict[str, str]]] = None,
    fontFamily: Optional[str] = None,
    landscape: Optional[bool] = None,
    tagged: Optional[bool] = None,
    outline: Optional[bool] = None,
    title: Optional[str] = None,
    author: Optional[str] = None
) -> str:
    """contentFormat: "markdown" ou "json" (lista de blocos tipados:
    heading/paragraph/list/table/image/code/blockquote/hr/pagebreak).
    Imagem local em qualquer um dos dois formatos e embutida automaticamente
    (nao depende do arquivo continuar no mesmo lugar depois).

    Parametros validos so para type="pdf":
    - css: substitui o estilo padrao por completo.
    - headerHtml/footerHtml: cabecalho/rodape repetidos em toda pagina
      (papel timbrado), via header_template/footer_template do Playwright.
      Suportam <span class="pageNumber">/"totalPages"/"date"/"title"/"url">.
    - fonts: lista de {"family", "path", "weight"?, "style"?} - fonte
      customizada (.ttf/.otf/.woff/.woff2) embutida no PDF via @font-face,
      funciona independente de estar instalada no sistema.
    - landscape, tagged (PDF acessivel/com estrutura), outline (sumario
      navegavel a partir dos headings).

    Parametro valido so para type="docx":
    - fontFamily: nome de fonte pro documento inteiro (precisa estar
      instalada em quem for abrir o arquivo - DOCX nao embute fonte).

    title/author: metadata do documento, valem pros dois tipos."""
    opts = {
        "path": path,
        "type": type,
        "content": content,
        "contentFormat": contentFormat,
        "templatePath": templatePath,
        "mergeFields": mergeFields,
        "css": css,
        "headerHtml": headerHtml,
        "footerHtml": footerHtml,
        "margins": margins,
        "fonts": fonts,
        "fontFamily": fontFamily,
        "landscape": landscape,
        "tagged": tagged,
        "outline": outline,
        "title": title,
        "author": author
    }
    return write_doc_service(opts)

@mcp.tool(name="write_sheet")
def write_sheet(
    path: str,
    data: List[Any],
    columns: Optional[Dict[str, str]] = None,
    freezePanes: Optional[bool] = None,
    append: Optional[bool] = None,
    delimiter: Optional[str] = None
) -> str:
    """delimiter vale so para .csv (default ","; use ";" para Excel PT-BR).
    .xlsx novo sai com cabecalho em negrito e largura de coluna automatica."""
    opts = {
        "path": path,
        "data": data,
        "columns": columns,
        "freezePanes": freezePanes,
        "append": append,
        "delimiter": delimiter
    }
    return write_sheet_service(opts)

@mcp.tool(name="render_slide")
def render_slide(
    html: str,
    css: str,
    format: str,
    aspectRatio: str,
    outputPath: str,
    fonts: Optional[List[Dict[str, str]]] = None,
    tagged: Optional[bool] = None,
    outline: Optional[bool] = None,
    title: Optional[str] = None,
    author: Optional[str] = None
) -> str:
    """Imagem local em html e embutida automaticamente. fonts: lista de
    {"family", "path", "weight"?, "style"?} pra fonte customizada via
    @font-face. tagged/title/author valem so quando format="pdf"."""
    opts = {
        "html": html,
        "css": css,
        "format": format,
        "aspectRatio": aspectRatio,
        "outputPath": outputPath,
        "fonts": fonts,
        "tagged": tagged,
        "outline": outline,
        "title": title,
        "author": author
    }
    return render_slide_service(opts)

@mcp.tool(name="render_page")
def render_page(
    html: str,
    css: str,
    outputPath: str,
    generateTOC: Optional[bool] = None,
    margins: Optional[Dict[str, str]] = None,
    fonts: Optional[List[Dict[str, str]]] = None,
    landscape: Optional[bool] = None,
    tagged: Optional[bool] = None,
    outline: Optional[bool] = None,
    title: Optional[str] = None,
    author: Optional[str] = None
) -> str:
    """Imagem local em html e embutida automaticamente. fonts: lista de
    {"family", "path", "weight"?, "style"?} pra fonte customizada via
    @font-face. tagged: PDF acessivel/com estrutura. outline: sumario
    navegavel a partir dos headings. title/author: metadata do PDF."""
    opts = {
        "html": html,
        "css": css,
        "outputPath": outputPath,
        "generateTOC": generateTOC,
        "margins": margins,
        "fonts": fonts,
        "landscape": landscape,
        "tagged": tagged,
        "outline": outline,
        "title": title,
        "author": author
    }
    return render_page_service(opts)

@mcp.tool(name="render_image")
def render_image(
    html: str,
    css: str,
    outputPath: str,
    width: int,
    height: int,
    format: Optional[str] = "png",
    transparentBackground: Optional[bool] = None,
    selector: Optional[str] = None,
    fonts: Optional[List[Dict[str, str]]] = None,
    scale: Optional[float] = None
) -> str:
    """Gera uma imagem (PNG/JPEG) a partir de HTML/CSS via screenshot -
    util pra logo/icone/banner standalone, nao so documentos. width/height
    definem o viewport; se selector for informado (seletor CSS), captura só
    esse elemento (ex: uma div .badge) em vez da pagina inteira -
    dimensiona a imagem exatamente pelo tamanho do elemento.
    transparentBackground remove o fundo (so vale pra format="png").
    scale (ex: 2) aumenta a resolucao (device pixel ratio) mantendo o
    layout logico em width x height. Imagem local referenciada em html e
    embutida automaticamente, igual write_doc/render_page/render_slide."""
    opts = {
        "html": html,
        "css": css,
        "outputPath": outputPath,
        "width": width,
        "height": height,
        "format": format,
        "transparentBackground": transparentBackground,
        "selector": selector,
        "fonts": fonts,
        "scale": scale
    }
    return render_image_service(opts)

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

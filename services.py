import asyncio
import os
import sys
import tempfile
import uuid
import zipfile
import csv
import re
import io
import difflib
import json
import base64
import binascii
import mimetypes
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from typing import Optional, List, Dict, Any

import docx
import pypdf
import openpyxl
import mammoth
import markdownify
import markdown as md_lib
from bs4 import BeautifulSoup, NavigableString, Comment
from playwright.async_api import async_playwright
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from pptx import Presentation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from security import validate_path, validate_write_path
from config import get_read_sheet_max_rows, get_scan_max_matches
from pptx_service import create_pptx_from_html


MAX_ARCHIVE_ENTRIES = 100_000
MAX_ARCHIVE_UNCOMPRESSED_BYTES = 200 * 1024 * 1024


REMOTE_URL_RE = re.compile(r"^https?://", re.IGNORECASE)

_FONT_FORMAT_BY_EXT = {
    ".ttf": "truetype",
    ".otf": "opentype",
    ".woff": "woff",
    ".woff2": "woff2",
}
_FONT_MIME_BY_FORMAT = {
    "truetype": "font/ttf",
    "opentype": "font/otf",
    "woff": "font/woff",
    "woff2": "font/woff2",
}


def _is_remote_url(src: str) -> bool:
    return bool(REMOTE_URL_RE.match(src or ""))


def _is_data_uri(src: str) -> bool:
    return (src or "").startswith("data:")


def _fetch_remote_bytes(url: str, timeout: int = 10) -> bytes:
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            return resp.read()
    except (urllib.error.URLError, ValueError, TimeoutError) as e:
        raise ValueError(f"Nao foi possivel baixar imagem remota: {url} ({e})")


def _guess_mime(path_or_url: str) -> str:
    mime, _ = mimetypes.guess_type(path_or_url)
    return mime or "application/octet-stream"


def _image_bytes_to_data_uri(data: bytes, mime: str) -> str:
    return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"


def _embed_local_images_in_html(html: str) -> str:
    """Playwright/Chromium nao carrega <img src="..."> de forma confiavel
    quando src e um path bare ou file:// via page.set_content() - so data
    URI funciona de forma consistente. URLs http(s) e data URIs ja
    funcionam e sao deixadas como estao."""
    if "<img" not in html:
        return html
    soup = BeautifulSoup(html, "html.parser")
    changed = False
    for img in soup.find_all("img"):
        src = img.get("src", "")
        if not src or _is_data_uri(src) or _is_remote_url(src):
            continue
        resolved = os.path.abspath(os.path.expanduser(src))
        if not os.path.isfile(resolved):
            raise ValueError(f"Imagem nao encontrada: {src}")
        with open(resolved, "rb") as f:
            data = f.read()
        img["src"] = _image_bytes_to_data_uri(data, _guess_mime(resolved))
        changed = True
    return str(soup) if changed else html


def _build_font_face_css(fonts: Optional[list]) -> str:
    if not fonts:
        return ""
    rules = []
    for spec in fonts:
        family = spec.get("family")
        path = spec.get("path")
        if not family or not path:
            raise ValueError("Cada item de 'fonts' precisa de 'family' e 'path'")
        resolved = os.path.abspath(os.path.expanduser(path))
        if not os.path.isfile(resolved):
            raise ValueError(f"Arquivo de fonte nao encontrado: {path}")
        ext = os.path.splitext(resolved)[1].lower()
        fmt = _FONT_FORMAT_BY_EXT.get(ext)
        if not fmt:
            raise ValueError(
                f"Formato de fonte nao suportado: {ext} (use .ttf, .otf, .woff ou .woff2)"
            )
        with open(resolved, "rb") as f:
            data = f.read()
        b64 = base64.b64encode(data).decode("ascii")
        weight = spec.get("weight", "normal")
        style = spec.get("style", "normal")
        mime = _FONT_MIME_BY_FORMAT[fmt]
        rules.append(
            f"@font-face {{ font-family: '{family}'; "
            f"src: url(data:{mime};base64,{b64}) format('{fmt}'); "
            f"font-weight: {weight}; font-style: {style}; }}"
        )
    return "\n".join(rules)


def odt_xml_to_plain(xml_bytes: bytes) -> str:
    root = ET.fromstring(xml_bytes)
    texts = []
    for elem in root.iter():
        if elem.text and elem.text.strip():
            texts.append(elem.text.strip())
    return "\n\n".join(texts)


def read_odt(path_resolved: str) -> str:
    with zipfile.ZipFile(path_resolved, "r") as z:
        try:
            content = z.read("content.xml")
        except KeyError:
            raise ValueError("content.xml nao encontrado no ODT.")
    return odt_xml_to_plain(content)


def read_ods_to_rows(
    ods_path: str, sheet_name: Optional[str] = None
) -> List[List[Any]]:
    ns = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }
    with zipfile.ZipFile(ods_path, "r") as z:
        content = z.read("content.xml")
    root = ET.fromstring(content)
    tables = root.findall(".//table:table", ns)
    if not tables:
        return []
    target_table = tables[0]
    if sheet_name:
        for t in tables:
            name_attr = f"{{{ns['table']}}}name"
            if t.attrib.get(name_attr) == sheet_name:
                target_table = t
                break
    rows = []
    for row_elem in target_table.findall("table:table-row", ns):
        row_repeat_attr = f"{{{ns['table']}}}number-rows-repeated"
        row_repeat = row_elem.attrib.get(row_repeat_attr)
        r_rep = int(row_repeat) if row_repeat and row_repeat.isdigit() else 1
        if r_rep > 1000:
            r_rep = 1000
        row = []
        cells = row_elem.findall("table:table-cell", ns)
        for cell in cells:
            repeat_attr = f"{{{ns['table']}}}number-columns-repeated"
            col_repeat = cell.attrib.get(repeat_attr)
            c_rep = int(col_repeat) if col_repeat and col_repeat.isdigit() else 1
            if c_rep > 1000:
                c_rep = 1000
            val_p = cell.find("text:p", ns)
            val = val_p.text if val_p is not None else ""
            for _ in range(c_rep):
                row.append(val)
        while row and row[-1] == "":
            row.pop()
        for _ in range(r_rep):
            rows.append(list(row))
    while rows and not any(cell != "" for cell in rows[-1]):
        rows.pop()
    return rows


def truncate_markdown(
    text: str, max_chars: Optional[int], preview_only: Optional[bool]
) -> tuple:
    preview_cap = 12000
    cap = (
        max_chars if max_chars is not None else (preview_cap if preview_only else None)
    )
    if cap is None or len(text) <= cap:
        return text, False
    return (
        f"{text[:cap]}\n\n_[… texto truncado; use maxChars ou desative previewOnly …]_",
        True,
    )


def read_doc_service(file_path: str, options: dict) -> dict:
    resolved = validate_path(file_path, True)
    ext = os.path.splitext(resolved)[1].lower()

    note = None
    if options.get("includeComments"):
        note = "includeComments: comentarios em Word/PDF podem nao ser extraidos integralmente nesta versao."

    if ext == ".pdf":
        reader = pypdf.PdfReader(resolved)
        pages_text = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                pages_text.append(t)
        md = (
            "\n".join(pages_text).strip()
            if pages_text
            else "(PDF sem texto extraivel.)"
        )
        md, truncated = truncate_markdown(
            md, options.get("maxChars"), options.get("previewOnly")
        )
        return {"markdown": md, "imagePaths": [], "note": note, "truncated": truncated}

    if ext == ".odt":
        text = read_odt(resolved)
        md = text if text else "(ODT vazio.)"
        md, truncated = truncate_markdown(
            md, options.get("maxChars"), options.get("previewOnly")
        )
        return {"markdown": md, "imagePaths": [], "note": note, "truncated": truncated}

    if ext == ".pptx":
        presentation = Presentation(resolved)
        sections = []
        for index, slide in enumerate(presentation.slides, start=1):
            lines = []
            for shape in slide.shapes:
                if getattr(shape, "has_table", False):
                    table = shape.table
                    if table.rows:
                        lines.append("| " + " | ".join(cell.text.replace("|", "\\|") for cell in table.rows[0].cells) + " |")
                        lines.append("| " + " | ".join("---" for _ in table.rows[0].cells) + " |")
                        for row in table.rows[1:]:
                            lines.append("| " + " | ".join(cell.text.replace("|", "\\|") for cell in row.cells) + " |")
                elif getattr(shape, "has_text_frame", False):
                    text = shape.text.strip()
                    if text:
                        lines.append(text)
            sections.append(f"## Slide {index}\n\n" + "\n\n".join(lines))
        md, truncated = truncate_markdown(
            "\n\n".join(sections) or "(PPTX sem texto.)",
            options.get("maxChars"), options.get("previewOnly")
        )
        return {"markdown": md, "imagePaths": [], "note": note, "truncated": truncated}

    if ext == ".docx":
        image_paths = []
        if not options.get("extractImages"):
            doc = docx.Document(resolved)
            md = "\n".join(p.text for p in doc.paragraphs).strip()
            md = md if md else "(DOCX sem texto.)"
            md, truncated = truncate_markdown(
                md, options.get("maxChars"), options.get("previewOnly")
            )
            return {
                "markdown": md,
                "imagePaths": [],
                "note": note,
                "truncated": truncated,
            }

        tmp_dir = os.path.join(tempfile.gettempdir(), f"docgen-img-{uuid.uuid4().hex}")
        os.makedirs(tmp_dir, exist_ok=True)

        def convert_image(image):
            with image.open() as image_bytes:
                ext_img = (
                    ".png"
                    if "png" in image.content_type
                    else ".jpg"
                    if "jpeg" in image.content_type
                    else ".bin"
                )
                name = f"image-{len(image_paths) + 1}{ext_img}"
                out = os.path.join(tmp_dir, name)
                with open(out, "wb") as f:
                    f.write(image_bytes.read())
                image_paths.append(out)
                return {"src": out}

        with open(resolved, "rb") as docx_file:
            result = mammoth.convert_to_html(
                docx_file, convert_image=mammoth.images.inline(convert_image)
            )
            html = result.value

        md = markdownify.markdownify(html, heading_style="ATX").strip()
        md = md if md else "(DOCX sem texto.)"
        md, truncated = truncate_markdown(
            md, options.get("maxChars"), options.get("previewOnly")
        )
        return {
            "markdown": md,
            "imagePaths": image_paths,
            "note": note,
            "truncated": truncated,
        }


def parse_range(range_str: str, max_row: int, max_col: int):
    if not range_str or not range_str.strip():
        return None
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", range_str.strip(), re.IGNORECASE)
    if not m:
        return None

    def col_to_idx(letters):
        n = 0
        for ch in letters.upper():
            n = n * 26 + (ord(ch) - 64)
        return n - 1

    c0 = col_to_idx(m.group(1))
    r0 = int(m.group(2)) - 1
    c1 = col_to_idx(m.group(3))
    r1 = int(m.group(4)) - 1
    if r0 < 0 or r1 < 0 or c0 < 0 or c1 < 0 or r0 > r1 or c0 > c1:
        return None
    return {
        "r0": max(0, r0),
        "r1": min(max_row - 1, r1),
        "c0": max(0, c0),
        "c1": min(max_col - 1, c1),
    }


def read_sheet_service(file_path: str, options: dict) -> dict:
    resolved = validate_path(file_path, True)
    ext = os.path.splitext(resolved)[1].lower()

    if ext == ".csv":
        with open(resolved, mode="r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
    elif ext == ".ods":
        rows = read_ods_to_rows(resolved, options.get("sheetName"))
    elif ext in {".xlsx", ".xlsm"}:
        wb = openpyxl.load_workbook(resolved, read_only=True, data_only=True)
        name = options.get("sheetName")
        if not name or name not in wb.sheetnames:
            name = wb.sheetnames[0]
        ws = wb[name]
        rows = [[cell for cell in row] for row in ws.iter_rows(values_only=True)]
    else:
        raise ValueError(f"Extensao nao suportada para read_sheet: {ext}. Use .xlsx, .xlsm, .ods ou .csv.")

    if not rows:
        return {
            "asJson": options.get("asJson"),
            "markdown": "",
            "jsonRows": [],
            "truncated": False,
            "rowLimit": 0,
        }

    max_row = len(rows)
    max_col = max(len(r) for r in rows) if rows else 0
    for r in rows:
        if len(r) < max_col:
            r.extend([""] * (max_col - len(r)))

    range_str = options.get("range")
    if range_str and range_str.strip():
        bounds = parse_range(range_str, max_row, max_col)
        if not bounds:
            raise ValueError(
                f'range invalido: "{range_str}". Use o formato Excel com colunas e linhas, ex.: A1:D10 ou B2:B50.'
            )
        sliced_rows = []
        for r_idx in range(bounds["r0"], bounds["r1"] + 1):
            row = rows[r_idx]
            sliced_rows.append(row[bounds["c0"] : bounds["c1"] + 1])
        rows = sliced_rows

    env_cap = get_read_sheet_max_rows()
    max_rows = options.get("maxRows")
    cap = (
        min(max_rows, env_cap)
        if (max_rows is not None and max_rows > 0)
        else (min(500, env_cap) if options.get("previewOnly") else env_cap)
    )

    max_rows_total = 1 + cap
    truncated = len(rows) > max_rows_total
    sliced = rows[:max_rows_total] if truncated else rows

    if options.get("asJson"):
        headers = [str(h).strip() or f"col_{i}" for i, h in enumerate(sliced[0])]
        out_rows = []
        for r in sliced[1:]:
            obj = {}
            for idx, h in enumerate(headers):
                val = r[idx] if idx < len(r) else ""
                obj[h] = val
            out_rows.append(obj)
        return {
            "asJson": True,
            "jsonRows": out_rows,
            "truncated": truncated,
            "rowLimit": cap,
        }

    esc = lambda c: (
        str(c if c is not None else "").replace("|", "\\|").replace("\n", " ")
    )
    header_line = "| " + " | ".join(esc(h) for h in sliced[0]) + " |"
    sep_line = "| " + " | ".join("---" for _ in sliced[0]) + " |"
    body_lines = ["| " + " | ".join(esc(val) for val in r) + " |" for r in sliced[1:]]
    md = header_line + "\n" + sep_line + "\n" + "\n".join(body_lines)
    if truncated:
        md += f"\n\n_[Saida truncada: no maximo {cap} linhas de dados. Ajuste maxRows ou previewOnly._"
    return {"asJson": False, "markdown": md, "truncated": truncated, "rowLimit": cap}


def glob_to_regex(pattern: str) -> re.Pattern:
    p = pattern
    for ch in [".", "+", "^", "$", "{", "}", "(", ")", "|", "[", "]", "\\"]:
        p = p.replace(ch, "\\" + ch)
    p = p.replace("**", "{{GLOBSTAR}}")
    p = p.replace("*", "[^/]*")
    p = p.replace("?", ".")
    p = p.replace("{{GLOBSTAR}}", ".*")
    return re.compile(f"^{p}$", re.IGNORECASE)


def read_archive_service(file_path: str, pattern: Optional[str] = None) -> str:
    resolved = validate_path(file_path, True)
    ext = os.path.splitext(resolved)[1].lower()
    if ext != ".zip":
        raise ValueError(f"read_archive suporta apenas .zip. Recebido: {ext}")

    with zipfile.ZipFile(resolved, "r") as z:
        infos = z.infolist()
        if len(infos) > MAX_ARCHIVE_ENTRIES:
            raise ValueError(f"read_archive: ZIP excede o limite de {MAX_ARCHIVE_ENTRIES} entradas.")
        total_size = sum(info.file_size for info in infos)
        if total_size > MAX_ARCHIVE_UNCOMPRESSED_BYTES:
            raise ValueError("read_archive: conteúdo descompactado do ZIP excede o limite de 200 MB.")
        names = [info.filename for info in infos]

    if pattern:
        pat = pattern.strip()
        rx = glob_to_regex(pat)
        names = [n for n in names if rx.match(n.replace("\\", "/"))]

    root = {}
    for name in sorted(names):
        parts = [p for p in name.replace("\\", "/").split("/") if p]
        cur = root
        for idx, p in enumerate(parts):
            if idx == len(parts) - 1 and not name.endswith("/"):
                cur[p] = "(file)"
            else:
                if p not in cur or not isinstance(cur[p], dict):
                    cur[p] = {}
                cur = cur[p]

    lines = []

    def walk(obj: dict, prefix=""):
        for k in sorted(obj.keys()):
            v = obj[k]
            is_file = v == "(file)"
            lines.append(f"{prefix}{k}" + ("" if is_file else "/"))
            if not is_file and isinstance(v, dict):
                walk(v, f"{prefix}{k}/")

    walk(root)
    tree = "\n".join(lines) if lines else "(vazio)"
    header = f"Arquivo: {resolved}\nEntradas{f' (filtro: {pattern})' if pattern else ''}: {len(names)}\n\n"
    return header + tree


class _StrikethroughExtension(md_lib.Extension):
    def extendMarkdown(self, md):
        from markdown.inlinepatterns import SimpleTagPattern

        md.inlinePatterns.register(SimpleTagPattern(r"(~~)(.*?)\2", "del"), "del", 65)


def _md_to_html(md_content: str) -> str:
    lines = md_content.split("\n")
    out = []
    n = len(lines)
    prev_list_type = None
    i = 0
    while i < n:
        line = lines[i]
        stripped = line.strip()
        out.append(line.rstrip())
        is_ol = bool(re.match(r"^\s*\d+\.\s+", stripped))
        is_ul = bool(re.match(r"^\s*[-*+]\s+", stripped))
        is_blank = stripped == ""

        if is_blank:
            j = i + 1
            while j < n and lines[j].strip() == "":
                j += 1
            if j < n:
                next_line = lines[j].strip()
                next_ol = bool(re.match(r"^\s*\d+\.\s+", next_line))
                next_ul = bool(re.match(r"^\s*[-*+]\s+", next_line))
                if (prev_list_type == "ol" and next_ul) or (
                    prev_list_type == "ul" and next_ol
                ):
                    out.append("<!-- sep -->")
        elif is_ol:
            prev_list_type = "ol"
        elif is_ul:
            prev_list_type = "ul"
        else:
            prev_list_type = None
        i += 1

    return md_lib.markdown(
        "\n".join(out), extensions=["extra", _StrikethroughExtension()]
    )


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _inline_md_to_html(text: str) -> str:
    html_out = md_lib.markdown(text or "", extensions=["extra", _StrikethroughExtension()])
    if html_out.startswith("<p>") and html_out.endswith("</p>") and "<p>" not in html_out[3:-4]:
        return html_out[3:-4]
    return html_out


def _runs_to_html(runs: list) -> str:
    parts = []
    for r in runs:
        text = _esc(r.get("text", ""))
        if r.get("code"):
            text = f"<code>{text}</code>"
        if r.get("bold"):
            text = f"<strong>{text}</strong>"
        if r.get("italic"):
            text = f"<em>{text}</em>"
        if r.get("link"):
            text = f'<a href="{_esc(r["link"])}">{text}</a>'
        parts.append(text)
    return "".join(parts)


def _block_text_to_html(block: dict) -> str:
    if block.get("runs"):
        return _runs_to_html(block["runs"])
    return _inline_md_to_html(block.get("text", ""))


def _cell_to_blocks(cell) -> list:
    if isinstance(cell, str):
        return [{"type": "paragraph", "text": cell}]
    if isinstance(cell, list):
        return cell
    raise ValueError(f"Celula de tabela invalida: {cell!r} (use string ou lista de blocos)")


def _list_item_to_blocks(item) -> list:
    if isinstance(item, str):
        return [{"type": "paragraph", "text": item}]
    if isinstance(item, list):
        return item
    raise ValueError(f"Item de lista invalido: {item!r} (use string ou lista de blocos)")


def _block_to_html(block) -> str:
    if not isinstance(block, dict) or "type" not in block:
        raise ValueError(f"Bloco JSON invalido, esperado objeto com 'type': {block!r}")
    btype = block["type"]

    if btype == "heading":
        level = block.get("level", 1)
        if not isinstance(level, int) or not (1 <= level <= 6):
            raise ValueError(f"heading.level invalido: {level!r} (use 1-6)")
        return f"<h{level}>{_block_text_to_html(block)}</h{level}>"

    if btype == "paragraph":
        return f"<p>{_block_text_to_html(block)}</p>"

    if btype == "list":
        style = block.get("style", "bullet")
        tag = "ol" if style == "number" else "ul"
        items = block.get("items", [])
        items_html = "".join(
            f"<li>{json_blocks_to_html(_list_item_to_blocks(item))}</li>" for item in items
        )
        return f"<{tag}>{items_html}</{tag}>"

    if btype == "table":
        headers = block.get("headers")
        rows = block.get("rows", [])
        parts = ["<table>"]
        if headers:
            head_cells = "".join(f"<th>{_esc(str(h))}</th>" for h in headers)
            parts.append(f"<tr>{head_cells}</tr>")
        for row in rows:
            cells_html = "".join(
                f"<td>{json_blocks_to_html(_cell_to_blocks(cell))}</td>" for cell in row
            )
            parts.append(f"<tr>{cells_html}</tr>")
        parts.append("</table>")
        return "".join(parts)

    if btype == "image":
        src = block.get("src")
        if not src:
            raise ValueError("image.src e obrigatorio")
        alt = _esc(block.get("alt", "Imagem"))
        width_attr = f' width="{_esc(str(block["width"]))}"' if block.get("width") else ""
        return f'<img src="{_esc(src)}" alt="{alt}"{width_attr}/>'

    if btype == "code":
        text = _esc(block.get("text", ""))
        return f"<pre><code>{text}</code></pre>"

    if btype == "blockquote":
        inner = json_blocks_to_html(block.get("blocks", []))
        return f"<blockquote>{inner}</blockquote>"

    if btype == "hr":
        return "<hr/>"

    if btype == "pagebreak":
        return '<div class="docgen-pagebreak" style="page-break-after: always;"></div>'

    raise ValueError(f"type de bloco desconhecido: {btype!r}")


def json_blocks_to_html(blocks: list) -> str:
    if not isinstance(blocks, list):
        raise ValueError("content JSON deve ser uma lista de blocos no nivel raiz")
    return "".join(_block_to_html(b) for b in blocks)


def _content_to_html(content: str, fmt: str) -> str:
    if fmt == "markdown":
        return _md_to_html(content)
    if fmt == "json":
        try:
            blocks = json.loads(content)
        except json.JSONDecodeError as e:
            raise ValueError(f"content JSON invalido: {e}")
        return json_blocks_to_html(blocks)
    raise ValueError(f"contentFormat invalido: {fmt!r} (use 'markdown' ou 'json')")


def _add_inline_runs(paragraph, element):
    for child in element.children:
        if isinstance(child, Comment):
            continue
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                paragraph.add_run(text)
            continue
        tag = child.name
        if tag is None:
            continue
        if tag == "br":
            paragraph.add_run("\n")
            continue
        text = child.get_text()
        if not text:
            continue
        run = paragraph.add_run(text)
        if tag in ("strong", "b"):
            run.bold = True
        elif tag in ("em", "i"):
            run.italic = True
        elif tag in ("del", "s"):
            run.font.strike = True
        elif tag in ("code",):
            run.font.name = "Consolas"
            run.font.size = docx.shared.Pt(9)
        elif tag in ("u", "ins"):
            run.underline = True
        elif tag == "a":
            href = child.get("href", "")
            run.font.color.rgb = docx.shared.RGBColor(0x05, 0x63, 0xC1)
            run.underline = True
            if href:
                try:
                    from docx.opc.constants import RELATIONSHIP_TYPE as RT

                    paragraph.part.relate_to(href, RT.HYPERLINK, is_external=True)
                except Exception:
                    pass


_BLOCK_TAGS = {
    "p", "ul", "ol", "table", "pre", "blockquote", "hr", "img",
    "h1", "h2", "h3", "h4", "h5", "h6", "div",
}


def _is_pagebreak_marker(tag_elem) -> bool:
    classes = tag_elem.get("class") or []
    return "docgen-pagebreak" in classes


def _cell_has_block_children(cell_elem) -> bool:
    return any(
        getattr(c, "name", None) in _BLOCK_TAGS and not (
            getattr(c, "name", None) == "div" and _is_pagebreak_marker(c)
        )
        for c in cell_elem.children
    )


def _add_heading(container, text: str, level: int):
    if hasattr(container, "add_heading"):
        container.add_heading(text, level=level)
        return
    p = container.add_paragraph(text)
    try:
        p.style = f"Heading {level}"
    except KeyError:
        for run in p.runs:
            run.bold = True


def _docx_image_width(width_attr: Optional[str]):
    if not width_attr:
        return docx.shared.Inches(4)
    digits = re.sub(r"[^0-9]", "", width_attr)
    if not digits:
        return docx.shared.Inches(4)
    return docx.shared.Inches(int(digits) / 96)  # 96px = 1in, HTML default DPI


def _add_docx_image(container, img_elem):
    src = img_elem.get("src", "")
    alt = img_elem.get("alt", "Imagem")
    width = _docx_image_width(img_elem.get("width"))

    image_source = None
    if _is_remote_url(src):
        try:
            image_source = io.BytesIO(_fetch_remote_bytes(src))
        except ValueError:
            image_source = None
    elif _is_data_uri(src):
        try:
            _, b64data = src.split(",", 1)
            image_source = io.BytesIO(base64.b64decode(b64data))
        except (ValueError, binascii.Error):
            image_source = None
    elif os.path.isfile(src):
        image_source = src

    if image_source is not None and hasattr(container, "add_picture"):
        try:
            container.add_picture(image_source, width=width)
            return
        except Exception:
            pass
    container.add_paragraph(f"[{alt}: {src}]")


def _html_table_to_docx(container, table_elem):
    rows = table_elem.find_all("tr")
    if not rows:
        return
    row_cells = [r.find_all(["th", "td"]) for r in rows]
    cols = max((len(cells) for cells in row_cells), default=0)
    if cols == 0:
        return
    if not hasattr(container, "add_table"):
        container.add_paragraph(table_elem.get_text(" ", strip=True))
        return
    table = container.add_table(rows=len(rows), cols=cols)
    table.style = "Table Grid"
    for ri, cells in enumerate(row_cells):
        for ci in range(cols):
            cell = table.rows[ri].cells[ci]
            p = cell.paragraphs[0]
            p.clear()
            if ci >= len(cells):
                continue
            _fill_docx_cell(cell, cells[ci])


def _fill_docx_cell(cell, cell_elem):
    if not _cell_has_block_children(cell_elem):
        _add_inline_runs(cell.paragraphs[0], cell_elem)
        return
    _html_block_to_docx(cell, cell_elem)
    first_p = cell.paragraphs[0]
    if first_p.text == "" and not first_p.runs and len(cell.paragraphs) > 1:
        first_p._element.getparent().remove(first_p._element)


def _html_block_to_docx(container, element):
    for child in element.children:
        if isinstance(child, Comment):
            continue
        if isinstance(child, NavigableString):
            text = str(child).strip()
            if text:
                container.add_paragraph(text)
            continue
        tag = child.name
        if tag is None:
            continue
        if tag in set("h1 h2 h3 h4 h5 h6".split()):
            level = int(tag[1])
            _add_heading(container, child.get_text().strip(), level=level)
        elif tag == "p":
            p = container.add_paragraph()
            _add_inline_runs(p, child)
        elif tag == "ul":
            for li in child.find_all("li", recursive=False):
                p = container.add_paragraph(style="List Bullet")
                _add_inline_runs(p, li)
        elif tag == "ol":
            for li in child.find_all("li", recursive=False):
                p = container.add_paragraph(style="List Number")
                _add_inline_runs(p, li)
        elif tag == "table":
            _html_table_to_docx(container, child)
            container.add_paragraph()
        elif tag == "pre":
            code_elem = child.find("code")
            code_text = code_elem.get_text() if code_elem else child.get_text()
            for line in code_text.split("\n"):
                p = container.add_paragraph()
                run = p.add_run(line)
                run.font.name = "Consolas"
                run.font.size = docx.shared.Pt(9)
                p.paragraph_format.space_before = docx.shared.Pt(1)
                p.paragraph_format.space_after = docx.shared.Pt(1)
        elif tag == "blockquote":
            for block_child in child.children:
                if getattr(block_child, "name", None) == "p":
                    p = container.add_paragraph()
                    p.paragraph_format.left_indent = docx.shared.Inches(0.5)
                    _add_inline_runs(p, block_child)
        elif tag == "hr":
            container.add_paragraph("─" * 50)
        elif tag == "img":
            _add_docx_image(container, child)
        elif tag == "div" and _is_pagebreak_marker(child):
            if hasattr(container, "add_page_break"):
                container.add_page_break()
        else:
            # unrecognized tag: recurse into it if it has block children,
            # otherwise fall back to its flattened text so nothing is
            # silently dropped.
            if child.find(True) is not None:
                _html_block_to_docx(container, child)
            else:
                text = child.get_text().strip()
                if text:
                    container.add_paragraph(text)


def html_to_docx_paragraphs(doc, html: str):
    soup = BeautifulSoup(html, "html.parser")
    _html_block_to_docx(doc, soup)


def _apply_docx_font_family(doc, font_family: str):
    doc.styles["Normal"].font.name = font_family
    for level in range(1, 10):
        style_name = f"Heading {level}"
        if style_name in doc.styles:
            doc.styles[style_name].font.name = font_family


def write_docx_from_content(
    content: str,
    fmt: str,
    out_path: str,
    font_family: Optional[str] = None,
    title: Optional[str] = None,
    author: Optional[str] = None,
):
    html = _content_to_html(content, fmt)
    doc = docx.Document()
    if font_family:
        _apply_docx_font_family(doc, font_family)
    html_to_docx_paragraphs(doc, html)
    if title:
        doc.core_properties.title = title
    if author:
        doc.core_properties.author = author
    doc.save(out_path)


def merge_docx_template(template_path: str, output_path: str, merge_fields: dict):
    with zipfile.ZipFile(template_path, "r") as jin:
        with zipfile.ZipFile(output_path, "w") as jout:
            for item in jin.infolist():
                data = jin.read(item.filename)
                if item.filename == "word/document.xml":
                    xml_content = data.decode("utf-8")
                    for k, v in merge_fields.items():
                        token = f"{{{{{k}}}}}"
                        xml_content = xml_content.replace(
                            token, str(v if v is not None else "")
                        )
                    data = xml_content.encode("utf-8")
                jout.writestr(item, data)


DEFAULT_PDF_CSS = "body{font-family:system-ui,sans-serif;font-size:11pt;line-height:1.5;margin:2cm}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px 10px;text-align:left}th{background:#f5f5f5}code{background:#f0f0f0;padding:1px 4px;border-radius:3px;font-family:Consolas,monospace;font-size:9.5pt}pre{background:#f5f5f5;padding:12px;border-radius:4px;overflow-x:auto}pre code{background:transparent;padding:0}blockquote{border-left:3px solid #ccc;margin-left:0;padding-left:16px;color:#555}img{max-width:100%}h1{font-size:20pt}h2{font-size:16pt}h3{font-size:13pt}hr{border:none;border-top:1px solid #ccc}"

DEFAULT_PDF_MARGIN = {"top": "20mm", "right": "15mm", "bottom": "20mm", "left": "15mm"}


def _set_pdf_metadata(out_path: str, title: Optional[str], author: Optional[str]):
    if not title and not author:
        return
    reader = pypdf.PdfReader(out_path)
    writer = pypdf.PdfWriter()
    writer.append(reader)
    meta = {}
    if title:
        meta["/Title"] = title
    if author:
        meta["/Author"] = author
    writer.add_metadata(meta)
    with open(out_path, "wb") as f:
        writer.write(f)


async def write_pdf_service(
    content: str,
    fmt: str,
    out_path: str,
    css: Optional[str] = None,
    header_html: Optional[str] = None,
    footer_html: Optional[str] = None,
    margins: Optional[dict] = None,
    fonts: Optional[list] = None,
    landscape: Optional[bool] = None,
    tagged: Optional[bool] = None,
    outline: Optional[bool] = None,
    title: Optional[str] = None,
    author: Optional[str] = None,
):
    html_body = _embed_local_images_in_html(_content_to_html(content, fmt))
    font_faces = _build_font_face_css(fonts)
    style = (font_faces + "\n" if font_faces else "") + (
        css if css is not None else DEFAULT_PDF_CSS
    )
    title_tag = f"<title>{_esc(title)}</title>" if title else ""
    full_html = f"<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'/>{title_tag}<style>{style}</style></head><body>{html_body}</body></html>"

    margin = dict(DEFAULT_PDF_MARGIN)
    if margins:
        margin.update({k: v for k, v in margins.items() if v})

    show_header_footer = bool(header_html or footer_html)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, channel="chrome")
        page = await browser.new_page()
        await page.set_content(full_html)
        await page.pdf(
            path=out_path,
            format="A4",
            print_background=True,
            margin=margin,
            display_header_footer=show_header_footer,
            header_template=header_html or "<span></span>",
            footer_template=footer_html or "<span></span>",
            landscape=bool(landscape),
            tagged=bool(tagged),
            outline=bool(outline),
        )
        await browser.close()

    _set_pdf_metadata(out_path, title, author)


async def write_doc_service(params: dict) -> str:
    out_path = validate_write_path(params["path"])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    fmt = params.get("contentFormat", "markdown")
    dtype = params["type"]

    if dtype == "docx":
        if params.get("templatePath"):
            tpl = validate_path(params["templatePath"], True)
            merge_fields = params.get("mergeFields", {})
            merge_docx_template(tpl, out_path, merge_fields)
            return (
                f"DOCX gravado em {out_path} (template + mergeFields)."
                if merge_fields
                else f"DOCX gravado em {out_path} (copia do template)."
            )
        write_docx_from_content(
            params["content"],
            fmt,
            out_path,
            font_family=params.get("fontFamily"),
            title=params.get("title"),
            author=params.get("author"),
        )
        return f"DOCX gravado em {out_path}."

    if dtype == "pdf":
        await write_pdf_service(
            params["content"],
            fmt,
            out_path,
            css=params.get("css"),
            header_html=params.get("headerHtml"),
            footer_html=params.get("footerHtml"),
            margins=params.get("margins"),
            fonts=params.get("fonts"),
            landscape=params.get("landscape"),
            tagged=params.get("tagged"),
            outline=params.get("outline"),
            title=params.get("title"),
            author=params.get("author"),
        )
        return f"PDF gravado em {out_path}."

    raise ValueError(f"Tipo nao suportado: {dtype}")


def write_sheet_service(params: dict) -> str:
    out_path = validate_write_path(params["path"])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    ext = os.path.splitext(out_path)[1].lower()
    data = params["data"]
    columns = params.get("columns") or {}

    if not isinstance(data, list):
        raise ValueError("write_sheet: data deve ser uma lista.")

    if not data:
        if ext == ".csv":
            with open(out_path, "w", encoding="utf-8-sig") as f:
                f.write("")
            return f"CSV vazio gravado em {out_path}."
        elif ext == ".xlsx":
            wb = openpyxl.Workbook()
            wb.save(out_path)
            return f"Planilha vazia gravada em {out_path}."
        else:
            raise ValueError("write_sheet: use extensao .xlsx ou .csv.")

    first = data[0]
    if isinstance(first, dict):
        keys = list(first.keys())
    elif isinstance(first, list):
        keys = [f"col{i + 1}" for i in range(len(first))]
    else:
        keys = ["value"]

    headers = [columns.get(k, k) for k in keys]

    row_values = []
    for r in data:
        if isinstance(r, dict):
            row_values.append([r.get(k) for k in keys])
        elif isinstance(r, list):
            row_values.append(r)
        else:
            row_values.append([r])

    if ext == ".csv":
        delimiter = params.get("delimiter") or ","
        if not isinstance(delimiter, str) or len(delimiter) != 1:
            raise ValueError("write_sheet: delimiter deve conter exatamente um caractere.")
        with open(out_path, mode="w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=delimiter)
            writer.writerow(headers)
            writer.writerows(row_values)
        return f"CSV gravado em {out_path} ({len(row_values)} linha(s))."

    if ext == ".xlsx":
        if params.get("append") and os.path.exists(out_path):
            wb = openpyxl.load_workbook(out_path)
            ws = wb.worksheets[0]
            for rv in row_values:
                ws.append(rv)
            wb.save(out_path)
            return f"Linhas anexadas em {out_path} ({len(row_values)} linha(s))."

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(headers)
        header_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        for rv in row_values:
            ws.append(rv)
        if params.get("freezePanes"):
            ws.freeze_panes = "A2"
        _autofit_columns(ws, headers, row_values)
        wb.save(out_path)
        return f"Planilha gravada em {out_path} ({len(row_values)} linha(s) de dados + cabecalho)."

    raise ValueError("write_sheet: use extensao .xlsx ou .csv.")


def _autofit_columns(ws, headers: list, row_values: list, max_width: int = 60):
    widths = [len(str(h)) for h in headers]
    for row in row_values:
        for i, val in enumerate(row):
            if i >= len(widths):
                continue
            widths[i] = max(widths[i], len(str(val)) if val is not None else 0)
    for i, width in enumerate(widths, start=1):
        letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[letter].width = min(max(width, 8) + 2, max_width)


async def render_slide_service(params: dict) -> str:
    out_path = validate_write_path(params["outputPath"])
    if params["format"].lower() != "pptx":
        raise ValueError("render_slide: format deve ser 'pptx'.")
    unsupported = [name for name in ("fonts", "tagged", "outline") if params.get(name)]
    if unsupported:
        raise ValueError(f"render_slide: parâmetros não suportados para PPTX: {', '.join(unsupported)}.")
    aspect_ratio = params["aspectRatio"]
    output = await asyncio.to_thread(
        create_pptx_from_html,
        params["html"],
        out_path,
        params.get("title"),
        params.get("author"),
        aspect_ratio,
        params.get("css") or "",
        params.get("footerHtml"),
        params.get("assetBasePath"),
    )
    return f"Slides exportados para PPTX: {output}"


async def render_page_service(params: dict) -> str:
    out_path = validate_write_path(params["outputPath"])
    toc_css = ""
    body_extra = ""
    if params.get("generateTOC"):
        toc_css = """
          #docgen-toc { page-break-after: always; font-family: system-ui, sans-serif; }
          #docgen-toc h2 { font-size: 14pt; }
          #docgen-toc nav { margin-top: 12px; }
          #docgen-toc a { color: inherit; text-decoration: none; }
        """
        body_extra = """
          <div id="docgen-toc"><h2>Indice</h2><nav id="toc-nav"></nav></div>
          <script>
            document.addEventListener('DOMContentLoaded', () => {
              const nav = document.getElementById('toc-nav');
              const hs = document.querySelectorAll('h1, h2, h3');
              hs.forEach((h, i) => {
                if (!h.id) h.id = 'heading-' + i;
                const a = document.createElement('a');
                a.href = '#' + h.id;
                a.textContent = h.textContent || '';
                const p = document.createElement('p');
                p.appendChild(a);
                nav.appendChild(p);
              });
            });
          </script>
        """

    margin = params.get("margins") or {}
    top = margin.get("top", "20mm")
    right = margin.get("right", "15mm")
    bottom = margin.get("bottom", "20mm")
    left = margin.get("left", "15mm")

    body_html = _embed_local_images_in_html(params["html"])
    font_faces = _build_font_face_css(params.get("fonts"))
    style = (font_faces + "\n" if font_faces else "") + toc_css + "\n" + params["css"]
    title = params.get("title")
    title_tag = f"<title>{_esc(title)}</title>" if title else ""
    full_html = f"<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'/>{title_tag}<style>{style}</style></head><body>{body_html + body_extra}</body></html>"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, channel="chrome")
        page = await browser.new_page()
        await page.set_content(full_html)
        await page.pdf(
            path=out_path,
            format="A4",
            print_background=True,
            margin={"top": top, "right": right, "bottom": bottom, "left": left},
            display_header_footer=False,
            landscape=bool(params.get("landscape")),
            tagged=bool(params.get("tagged")),
            outline=bool(params.get("outline")),
        )
        await browser.close()

    _set_pdf_metadata(out_path, title, params.get("author"))
    return f"Documento PDF gerado: {out_path}"


async def render_image_service(params: dict) -> str:
    out_path = validate_write_path(params["outputPath"])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    width = params["width"]
    height = params["height"]

    fmt = (params.get("format") or "png").lower()
    if fmt == "jpg":
        fmt = "jpeg"
    if fmt not in ("png", "jpeg"):
        raise ValueError(f"format invalido: {fmt!r} (use 'png' ou 'jpeg')")

    transparent = bool(params.get("transparentBackground")) and fmt == "png"

    body_html = _embed_local_images_in_html(params["html"])
    font_faces = _build_font_face_css(params.get("fonts"))
    bg_reset = "background: transparent;" if transparent else ""
    style = (font_faces + "\n" if font_faces else "") + params.get("css", "")
    full_html = f"<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'/><style>html,body{{margin:0;padding:0;{bg_reset}}}{style}</style></head><body>{body_html}</body></html>"

    scale = params.get("scale") or 1

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, channel="chrome")
        page = await browser.new_page(
            viewport={"width": width, "height": height},
            device_scale_factor=scale,
        )
        await page.set_content(full_html)
        await page.wait_for_timeout(50)

        selector = params.get("selector")
        target = await page.query_selector(selector) if selector else None
        if selector and target is None:
            await browser.close()
            raise ValueError(f"selector nao encontrou nenhum elemento: {selector}")

        shot_kwargs = {"path": out_path, "type": fmt}
        if fmt == "png":
            shot_kwargs["omit_background"] = transparent
        if target is not None:
            await target.screenshot(**shot_kwargs)
        else:
            await page.screenshot(full_page=True, **shot_kwargs)
        await browser.close()

    return f"Imagem gerada: {out_path}"


def patch_doc_service(params: dict) -> str:
    action = params["action"]

    if action == "merge":
        out_path = validate_write_path(params["path"])
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        sources = [
            validate_path(s, True) for s in params.get("payload", {}).get("sources", [])
        ]
        if not sources:
            raise ValueError("merge: informe payload.sources (PDFs na ordem desejada).")

        merger = pypdf.PdfMerger()
        for src in sources:
            if os.path.splitext(src)[1].lower() != ".pdf":
                raise ValueError(
                    f"merge: apenas PDF suportado. Arquivo invalido: {src}"
                )
            merger.append(src)
        merger.write(out_path)
        merger.close()
        return f"PDF mesclado gravado em {out_path} ({len(sources)} arquivo(s))."

    resolved = validate_path(params["path"], True)
    out_path = validate_write_path(resolved)
    ext = os.path.splitext(resolved)[1].lower()

    if action == "split":
        if ext != ".pdf":
            raise ValueError("split: apenas arquivo .pdf na entrada.")
        p = params.get("payload") or {}
        out_dir = p.get("outputDir")
        if out_dir:
            out_dir = validate_write_path(out_dir)
        else:
            base = os.path.splitext(resolved)[0]
            out_dir = os.path.abspath(f"{base}_split")
        os.makedirs(out_dir, exist_ok=True)

        reader = pypdf.PdfReader(resolved)
        count = len(reader.pages)
        for i, page in enumerate(reader.pages):
            writer = pypdf.PdfWriter()
            writer.add_page(page)
            part_path = os.path.join(out_dir, f"page-{i + 1:03d}.pdf")
            with open(part_path, "wb") as f:
                writer.write(f)
        return f"PDF dividido em {count} arquivo(s) em {out_dir}."

    if action == "watermark":
        if ext != ".pdf":
            raise ValueError("watermark: apenas .pdf suportado nesta versao.")
        p = params.get("payload") or {}
        text = p.get("text", "CONFIDENCIAL")
        opacity = p.get("opacity", 0.15)

        reader = pypdf.PdfReader(resolved)
        writer = pypdf.PdfWriter()

        for page in reader.pages:
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(width, height))
            can.setFillColor(Color(0.6, 0.6, 0.6, alpha=opacity))
            can.setFont("Helvetica-Bold", min(width, height) / 18)
            can.drawCentredString(width / 2, height / 2, text)
            can.save()

            packet.seek(0)
            watermark_reader = pypdf.PdfReader(packet)
            page.merge_page(watermark_reader.pages[0])
            writer.add_page(page)

        with open(out_path, "wb") as f:
            writer.write(f)
        return f"Marca d'agua aplicada em {out_path}."

    if action == "replace_text":
        p = params.get("payload") or {}
        reps = p.get("replacements", [])
        if not reps:
            raise ValueError("replace_text: informe payload.replacements.")

        if ext == ".docx":
            temp_path = out_path + ".tmp"
            try:
                with zipfile.ZipFile(resolved, "r") as jin:
                    with zipfile.ZipFile(temp_path, "w") as jout:
                        for item in jin.infolist():
                            data = jin.read(item.filename)
                            if item.filename == "word/document.xml":
                                xml_content = data.decode("utf-8")
                                for replacement in reps:
                                    if not isinstance(replacement, dict) or "from" not in replacement or "to" not in replacement:
                                        raise ValueError("replace_text: cada substituição precisa de from e to.")
                                    xml_content = xml_content.replace(str(replacement["from"]), str(replacement["to"]))
                                data = xml_content.encode("utf-8")
                            jout.writestr(item, data)
                os.replace(temp_path, out_path)
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            return f"Substituicoes aplicadas no DOCX em {out_path}."

        if ext == ".pdf":
            raise ValueError("replace_text em PDF nao e suportado. Exporte para DOCX.")

        raise ValueError(f"replace_text: extensao nao suportada: {ext}")

    raise ValueError(f"Acao desconhecida: {action}")


def patch_sheet_service(file_path: str, updates: list) -> str:
    resolved = validate_path(file_path, True)
    out_path = validate_write_path(resolved)
    ext = os.path.splitext(resolved)[1].lower()
    if ext != ".xlsx":
        raise ValueError("patch_sheet suporta apenas .xlsx nesta versao.")
    if not updates:
        raise ValueError("Informe ao menos uma atualizacao em updates.")

    wb = openpyxl.load_workbook(resolved)
    ws = wb.worksheets[0]

    for u in updates:
        if not isinstance(u, dict) or not isinstance(u.get("cell"), str):
            raise ValueError("patch_sheet: cada atualização deve conter cell.")
        cell_ref = u["cell"].strip()
        if not re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]*", cell_ref):
            raise ValueError(f"patch_sheet: referência de célula inválida: {cell_ref!r}.")
        cell = ws[cell_ref]
        cell.value = u["value"]
        style = u.get("style")
        if style and isinstance(style, dict):
            if "font" in style:
                font_opts = style["font"]
                cell.font = Font(**font_opts)
            if "fill" in style:
                fill_opts = style["fill"]
                cell.fill = PatternFill(**fill_opts)
            if "alignment" in style:
                align_opts = style["alignment"]
                cell.alignment = Alignment(**align_opts)
            if "border" in style:
                border_opts = style["border"]
                sides = {}
                for direction in ["left", "right", "top", "bottom"]:
                    if direction in border_opts:
                        d_opts = border_opts[direction]
                        b_style = d_opts.get("style")
                        color_opt = d_opts.get("color")
                        color_str = None
                        if isinstance(color_opt, dict):
                            color_str = color_opt.get("argb") or color_opt.get("theme")
                        elif isinstance(color_opt, str):
                            color_str = color_opt
                        sides[direction] = Side(border_style=b_style, color=color_str)
                cell.border = Border(**sides)
            if "numFmt" in style:
                cell.number_format = style["numFmt"]

    wb.save(out_path)
    return f"Planilha atualizada: {out_path} ({len(updates)} celula(s))."


def scan_zip(zip_path: str, query_re: re.Pattern) -> List[str]:
    hits = []
    with zipfile.ZipFile(zip_path, "r") as z:
        for name in z.namelist():
            clean_name = name.replace("\\", "/")
            if not clean_name.endswith("/") and query_re.search(clean_name):
                hits.append(f"{zip_path}::{clean_name}")
    return hits


def read_file_snippet(file_path: str, max_bytes: int = 8000) -> str:
    with open(file_path, "rb") as f:
        data = f.read(max_bytes)
        return data.decode("utf-8", errors="ignore")


def scan_dir_service(
    root_path: str, query: str, recursive: bool, max_matches: Optional[int] = None
) -> dict:
    resolved = validate_path(root_path, True)
    try:
        re_flags = re.IGNORECASE | re.MULTILINE | re.DOTALL
        query_re = re.compile(query, re_flags)
    except Exception:
        raise ValueError("query nao e uma expressao regular valida.")

    st = os.stat(resolved)
    cap = (
        max_matches
        if (max_matches is not None and max_matches > 0)
        else get_scan_max_matches()
    )

    def trim(arr: list) -> tuple:
        if len(arr) <= cap:
            return arr, False
        return arr[:cap], True

    if os.path.isfile(resolved) and resolved.lower().endswith(".zip"):
        hits = scan_zip(resolved, query_re)
        trimmed, truncated = trim(hits)
        return {"matches": trimmed, "truncated": truncated}

    if os.path.isfile(resolved):
        text = read_file_snippet(resolved)
        lines = text.split("\n")
        hits = []
        for i, line in enumerate(lines):
            if query_re.search(line):
                hits.append(f"{resolved}:{i + 1}:{line[:200]}")
        trimmed, truncated = trim(hits)
        return {"matches": trimmed, "truncated": truncated}

    matches = []
    for root, dirs, files in os.walk(resolved):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d != "__pycache__"]
        if not recursive and root != resolved:
            continue
        for file in files:
            if file.startswith("."):
                continue
            full_path = os.path.join(root, file)
            base = os.path.basename(full_path)
            if query_re.search(base):
                matches.append(full_path)
                continue
            if full_path.lower().endswith(".zip"):
                matches.extend(scan_zip(full_path, query_re))
                continue
            try:
                text = read_file_snippet(full_path, 256 * 1024)
                if query_re.search(text):
                    matches.append(full_path)
            except Exception:
                pass
        if not recursive:
            break

    trimmed, truncated = trim(matches)
    return {"matches": trimmed, "truncated": truncated}


def diff_file_service(path_a: str, path_b: str, mode: str) -> str:
    a = validate_path(path_a, True)
    b = validate_path(path_b, True)

    if mode == "data":
        ext_a = os.path.splitext(a)[1].lower()
        ext_b = os.path.splitext(b)[1].lower()
        if ext_a != ext_b:
            return "diff data: as extensoes devem coincidir para modo data."

        if ext_a in {".xlsx", ".xls", ".csv", ".ods"}:

            def read_data_to_str(p, ext):
                if ext == ".csv":
                    with open(p, mode="r", encoding="utf-8-sig") as f:
                        return json.dumps(list(csv.reader(f)), indent=2)
                elif ext == ".ods":
                    return json.dumps(read_ods_to_rows(p), indent=2)
                else:
                    wb = openpyxl.load_workbook(p, data_only=True)
                    res = {}
                    for name in wb.sheetnames:
                        ws = wb[name]
                        res[name] = [
                            [c for c in row] for row in ws.iter_rows(values_only=True)
                        ]
                    return json.dumps(res, indent=2)

            ja = read_data_to_str(a, ext_a)
            jb = read_data_to_str(b, ext_b)
            if ja == jb:
                return "Planilhas equivalentes (modo data)."

            diff = difflib.unified_diff(
                ja.splitlines(), jb.splitlines(), fromfile=a, tofile=b
            )
            out = list(diff)
            return "\n".join(out[:5000]) + (
                "\n... (truncado)" if len(out) > 5000 else ""
            )

        return "Modo data: apenas planilhas tabulares (.xlsx, .csv, .ods) suportadas."

    try:
        with open(a, "r", encoding="utf-8") as fa:
            lines_a = fa.readlines()
        with open(b, "r", encoding="utf-8") as fb:
            lines_b = fb.readlines()
    except Exception:
        return "Modo text: um dos arquivos nao e texto UTF-8 legivel; use modo data ou converta antes."

    diff = difflib.unified_diff(lines_a, lines_b, fromfile=a, tofile=b)
    out = list(diff)
    return "".join(out[:8000]) + ("\n... (truncado)" if len(out) > 8000 else "")


def bundle_zip_service(files: List[str], output_name: str) -> str:
    if not files:
        raise ValueError("Informe ao menos um arquivo em files.")
    resolved = [validate_path(f, True) for f in files]
    out_path = validate_write_path(output_name)
    names = [os.path.basename(path) for path in resolved]
    if len(names) != len(set(names)):
        raise ValueError("bundle_zip: arquivos com o mesmo nome não podem ser empacotados juntos.")

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in resolved:
            z.write(f, os.path.basename(f))

    return f"Arquivo ZIP criado: {out_path} ({len(files)} arquivo(s))."

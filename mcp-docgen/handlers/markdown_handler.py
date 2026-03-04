"""Convert Markdown text into structured sections for docx/pdf generation."""

import re
from pathlib import Path

from handlers.docx_handler import create_docx
from handlers.pdf_handler import create_pdf


def markdown_to_document(
    markdown: str,
    output_format: str = "docx",
    title: str | None = None,
    output_path: str = "",
) -> str:
    """Parse markdown and generate a docx or pdf document.

    Supported markdown elements:
      - # Headings (levels 1-4)
      - **bold**, *italic*
      - Tables (pipe-delimited)
      - ```code blocks```
      - - / * unordered lists
      - 1. ordered lists
      - ![caption](image_path) images
      - --- page breaks / horizontal rules

    Returns the absolute path of the generated file.
    """
    sections = _parse_markdown(markdown)
    doc_title = title or _extract_title(sections) or "Documento"

    if output_format.lower() == "pdf":
        return create_pdf(doc_title, sections, output_path)
    return create_docx(doc_title, sections, output_path)


def _extract_title(sections: list[dict]) -> str | None:
    for s in sections:
        if s.get("type") == "heading" and s.get("level", 99) <= 1:
            return s["text"]
    return None


def _parse_markdown(md: str) -> list[dict]:
    lines = md.split("\n")
    sections: list[dict] = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # --- code block ---
        if line.strip().startswith("```"):
            lang = line.strip().removeprefix("```").strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            sections.append({
                "type": "code_block",
                "code": "\n".join(code_lines),
                "language": lang or "text",
            })
            i += 1
            continue

        # --- page break / horizontal rule ---
        if re.match(r"^-{3,}$|^\*{3,}$|^_{3,}$", line.strip()):
            sections.append({"type": "page_break"})
            i += 1
            continue

        # --- heading ---
        heading_match = re.match(r"^(#{1,4})\s+(.+)$", line)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            sections.append({"type": "heading", "text": text, "level": level})
            i += 1
            continue

        # --- image ---
        img_match = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", line.strip())
        if img_match:
            caption = img_match.group(1)
            img_path = img_match.group(2)
            section = {"type": "image", "path": img_path, "width_inches": 5.0}
            if caption:
                section["caption"] = caption
            sections.append(section)
            i += 1
            continue

        # --- table ---
        if "|" in line and i + 1 < len(lines) and re.match(r"^\|?\s*[-:]+", lines[i + 1]):
            table_section = _parse_table(lines, i)
            if table_section:
                sections.append(table_section[0])
                i = table_section[1]
                continue

        # --- unordered list ---
        if re.match(r"^\s*[-*+]\s+", line):
            items, i = _collect_list_items(lines, i, ordered=False)
            sections.append({"type": "list", "items": items, "ordered": False})
            continue

        # --- ordered list ---
        if re.match(r"^\s*\d+\.\s+", line):
            items, i = _collect_list_items(lines, i, ordered=True)
            sections.append({"type": "list", "items": items, "ordered": True})
            continue

        # --- empty line ---
        if not line.strip():
            i += 1
            continue

        # --- paragraph (with inline formatting) ---
        text = _process_inline(line.strip())
        bold = text.startswith("**") and text.endswith("**")
        italic = text.startswith("*") and text.endswith("*") and not bold

        if bold:
            sections.append({"type": "paragraph", "text": text.strip("*").strip(), "bold": True})
        elif italic:
            sections.append({"type": "paragraph", "text": text.strip("*").strip(), "italic": True})
        else:
            sections.append({"type": "paragraph", "text": text})
        i += 1

    return sections


def _parse_table(lines: list[str], start: int) -> tuple[dict, int] | None:
    header_line = lines[start].strip().strip("|")
    headers = [h.strip() for h in header_line.split("|")]

    i = start + 2
    rows = []
    while i < len(lines):
        line = lines[i].strip()
        if not line or "|" not in line:
            break
        row_line = line.strip("|")
        cells = [c.strip() for c in row_line.split("|")]
        rows.append(cells)
        i += 1

    return ({"type": "table", "headers": headers, "rows": rows}, i)


def _collect_list_items(lines: list[str], start: int, ordered: bool) -> tuple[list[str], int]:
    items = []
    i = start
    pattern = r"^\s*\d+\.\s+" if ordered else r"^\s*[-*+]\s+"

    while i < len(lines):
        match = re.match(pattern, lines[i])
        if not match:
            break
        text = re.sub(pattern, "", lines[i]).strip()
        items.append(_process_inline(text))
        i += 1

    return items, i


def _process_inline(text: str) -> str:
    """Strip markdown inline markers for plain text output."""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"__(.+?)__", r"\1", text)
    text = re.sub(r"_(.+?)_", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    return text

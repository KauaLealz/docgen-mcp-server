from fpdf import FPDF
from pathlib import Path
from PIL import Image as PILImage
import pdfplumber
import io

from utils.security import validate_image_path
from utils.file_utils import generate_output_path

FONT_DIR = Path(__file__).parent.parent / "fonts"

MAX_IMAGE_WIDTH_PX = 2000
MAX_IMAGE_HEIGHT_PX = 2000


class DocGenPDF(FPDF):
    """PDF subclass with Unicode font support and consistent styling."""

    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=20)
        self._setup_fonts()

    def _setup_fonts(self):
        self.add_page()
        self.set_font("Helvetica", size=11)

    def header(self):
        pass

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Página {self.page_no()}/{{nb}}", align="C")
        self.set_text_color(0, 0, 0)


def create_pdf(title: str, sections: list[dict], output_path: str = "") -> str:
    """Generate a PDF document from structured sections.

    Same section schema as create_docx:
      heading, paragraph, table, image, code_block, list, page_break

    Returns the absolute path of the generated file.
    """
    pdf = DocGenPDF()
    pdf.alias_nb_pages()

    if title:
        pdf.set_font("Helvetica", "B", 20)
        pdf.cell(0, 15, title, ln=True, align="C")
        pdf.ln(8)

    for section in sections:
        _render_section(pdf, section)

    dest = generate_output_path(output_path)
    pdf.output(str(dest))
    return str(dest)


def read_pdf(file_path: str, pages: list[int] | None = None) -> dict:
    """Extract text, tables, and metadata from a PDF.

    Args:
        file_path: Path to the PDF.
        pages: Optional 0-based page indices to extract. None = all.

    Returns:
        {
            "pages": [{"page_number": int, "text": str, "tables": [...]}],
            "total_pages": int,
            "metadata": dict
        }
    """
    resolved = validate_path(file_path, must_exist=True)
    result_pages = []

    with pdfplumber.open(str(resolved)) as pdf:
        metadata = pdf.metadata or {}
        total = len(pdf.pages)

        target_indices = pages if pages else range(total)
        for idx in target_indices:
            if idx < 0 or idx >= total:
                continue
            page = pdf.pages[idx]
            text = page.extract_text() or ""
            tables_raw = page.extract_tables() or []
            tables = []
            for t in tables_raw:
                if t and len(t) > 1:
                    headers = [str(c) if c else "" for c in t[0]]
                    rows = [[str(c) if c else "" for c in row] for row in t[1:]]
                    tables.append({"headers": headers, "rows": rows})
                elif t and len(t) == 1:
                    tables.append({"headers": [str(c) if c else "" for c in t[0]], "rows": []})

            result_pages.append({
                "page_number": idx + 1,
                "text": text,
                "tables": tables,
            })

    safe_metadata = {}
    for k, v in metadata.items():
        try:
            safe_metadata[k] = str(v)
        except Exception:
            safe_metadata[k] = repr(v)

    return {
        "pages": result_pages,
        "total_pages": total,
        "metadata": safe_metadata,
    }


def _render_section(pdf: DocGenPDF, section: dict):
    section_type = section.get("type", "paragraph")

    if section_type == "heading":
        level = section.get("level", 1)
        sizes = {0: 20, 1: 16, 2: 14, 3: 12, 4: 11}
        size = sizes.get(level, 11)
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", size)
        pdf.multi_cell(0, size * 0.6, section.get("text", ""))
        pdf.ln(3)
        pdf.set_font("Helvetica", size=11)

    elif section_type == "paragraph":
        text = section.get("text", "")
        if section.get("bold"):
            pdf.set_font("Helvetica", "B", 11)
        elif section.get("italic"):
            pdf.set_font("Helvetica", "I", 11)
        else:
            pdf.set_font("Helvetica", size=11)
        pdf.multi_cell(0, 6, text)
        pdf.ln(2)
        pdf.set_font("Helvetica", size=11)

    elif section_type == "table":
        headers = section.get("headers", [])
        rows = section.get("rows", [])
        if not headers:
            return
        col_count = len(headers)
        page_w = pdf.w - pdf.l_margin - pdf.r_margin
        col_w = page_w / col_count

        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(220, 230, 241)
        for h in headers:
            pdf.cell(col_w, 7, str(h)[:40], border=1, fill=True)
        pdf.ln()

        pdf.set_font("Helvetica", size=9)
        pdf.set_fill_color(255, 255, 255)
        for row in rows:
            for j, val in enumerate(row):
                if j < col_count:
                    pdf.cell(col_w, 6, str(val)[:40], border=1)
            pdf.ln()
        pdf.ln(4)
        pdf.set_font("Helvetica", size=11)

    elif section_type == "image":
        image_path = validate_image_path(section["path"])
        width_inches = section.get("width_inches", 5.0)
        width_mm = width_inches * 25.4

        resized = _resize_if_needed(image_path)
        img_source = str(resized) if resized else str(image_path)

        available_w = pdf.w - pdf.l_margin - pdf.r_margin
        if width_mm > available_w:
            width_mm = available_w

        x = (pdf.w - width_mm) / 2
        pdf.image(img_source, x=x, w=width_mm)
        pdf.ln(3)

        caption = section.get("caption")
        if caption:
            pdf.set_font("Helvetica", "I", 9)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 5, caption, ln=True, align="C")
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Helvetica", size=11)
        pdf.ln(2)

    elif section_type == "code_block":
        code = section.get("code", "")
        pdf.set_font("Courier", size=8)
        pdf.set_fill_color(245, 245, 245)
        pdf.set_x(pdf.l_margin + 5)
        for line in code.split("\n"):
            pdf.cell(0, 4.5, "  " + line, ln=True, fill=True)
        pdf.ln(4)
        pdf.set_font("Helvetica", size=11)

    elif section_type == "list":
        items = section.get("items", [])
        ordered = section.get("ordered", False)
        pdf.set_font("Helvetica", size=11)
        for i, item in enumerate(items):
            bullet = f"{i+1}. " if ordered else "  - "
            pdf.cell(0, 6, f"{bullet}{item}", ln=True)
        pdf.ln(2)

    elif section_type == "page_break":
        pdf.add_page()


def _resize_if_needed(image_path: Path) -> Path | None:
    """Resize oversized images to avoid bloated PDFs. Returns new path or None."""
    try:
        with PILImage.open(image_path) as img:
            w, h = img.size
            if w <= MAX_IMAGE_WIDTH_PX and h <= MAX_IMAGE_HEIGHT_PX:
                return None
            ratio = min(MAX_IMAGE_WIDTH_PX / w, MAX_IMAGE_HEIGHT_PX / h)
            new_size = (int(w * ratio), int(h * ratio))
            resized = img.resize(new_size, PILImage.LANCZOS)
            tmp_path = image_path.parent / f"_resized_{image_path.name}"
            resized.save(tmp_path)
            return tmp_path
    except Exception:
        return None

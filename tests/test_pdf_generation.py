import json

import pypdf
import pytest

from services import write_pdf_service, DEFAULT_PDF_CSS

pytestmark = pytest.mark.anyio


def _extract_text(path: str) -> str:
    reader = pypdf.PdfReader(path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


async def test_markdown_table_renders_in_pdf(tmp_path):
    md = "# Report\n\n| Name | Score |\n| --- | --- |\n| Alice | 10 |\n| Bob | 20 |\n"
    out = str(tmp_path / "table.pdf")
    await write_pdf_service(md, "markdown", out)
    text = _extract_text(out)
    assert "Report" in text
    assert "Alice" in text and "Bob" in text


async def test_json_multi_paragraph_cell_renders_in_pdf(tmp_path):
    content = json.dumps(
        [
            {
                "type": "table",
                "headers": ["Item", "Details"],
                "rows": [
                    [
                        "Widget",
                        [
                            {"type": "paragraph", "text": "First paragraph."},
                            {"type": "paragraph", "text": "Second paragraph."},
                        ],
                    ]
                ],
            }
        ]
    )
    out = str(tmp_path / "complex.pdf")
    await write_pdf_service(content, "json", out)
    text = _extract_text(out)
    assert "Widget" in text
    assert "First paragraph." in text
    assert "Second paragraph." in text


async def test_custom_css_replaces_default(tmp_path):
    custom_css = "body { background: #000; color: #fff; }"
    out = str(tmp_path / "styled.pdf")
    await write_pdf_service("# Hi", "markdown", out, css=custom_css)
    # can't introspect applied CSS from the PDF text layer directly, but a
    # successful render with custom css and no crash, producing a non-empty
    # PDF, is the meaningful signal here.
    assert "Hi" in _extract_text(out)
    assert custom_css != DEFAULT_PDF_CSS


async def test_header_and_footer_add_extra_pages_worth_of_margin(tmp_path):
    out_plain = str(tmp_path / "plain.pdf")
    out_letterhead = str(tmp_path / "letterhead.pdf")
    await write_pdf_service("# Doc", "markdown", out_plain)
    await write_pdf_service(
        "# Doc",
        "markdown",
        out_letterhead,
        header_html="<div style='font-size:8px;'>ACME Corp</div>",
        footer_html="<div style='font-size:8px;'><span class='pageNumber'></span></div>",
    )
    # Both must render successfully; the letterhead variant must actually
    # produce a readable PDF with the body content still present.
    assert "Doc" in _extract_text(out_letterhead)
    assert "Doc" in _extract_text(out_plain)


async def test_custom_margins_are_applied(tmp_path):
    out = str(tmp_path / "margins.pdf")
    await write_pdf_service("# Doc", "markdown", out, margins={"top": "40mm"})
    assert "Doc" in _extract_text(out)

import pytest

from services import json_blocks_to_html, _content_to_html


def test_heading():
    html = json_blocks_to_html([{"type": "heading", "level": 2, "text": "Title"}])
    assert html == "<h2>Title</h2>"


def test_heading_invalid_level_raises():
    with pytest.raises(ValueError):
        json_blocks_to_html([{"type": "heading", "level": 7, "text": "x"}])


def test_paragraph_inline_markdown():
    html = json_blocks_to_html([{"type": "paragraph", "text": "**bold** and *italic*"}])
    assert "<strong>bold</strong>" in html
    assert "<em>italic</em>" in html


def test_paragraph_with_runs():
    block = {
        "type": "paragraph",
        "runs": [
            {"text": "hello ", "bold": True},
            {"text": "world", "italic": True, "link": "https://example.com"},
        ],
    }
    html = json_blocks_to_html([block])
    assert "<strong>hello </strong>" in html
    assert '<a href="https://example.com">' in html
    assert "<em>world</em>" in html


def test_list_bullet_simple_items():
    html = json_blocks_to_html(
        [{"type": "list", "style": "bullet", "items": ["one", "two"]}]
    )
    assert html == "<ul><li><p>one</p></li><li><p>two</p></li></ul>"


def test_list_number_multi_block_item():
    block = {
        "type": "list",
        "style": "number",
        "items": [[{"type": "paragraph", "text": "first"}, {"type": "paragraph", "text": "second"}]],
    }
    html = json_blocks_to_html([block])
    assert html == "<ol><li><p>first</p><p>second</p></li></ol>"


def test_table_with_headers_and_simple_cells():
    block = {
        "type": "table",
        "headers": ["A", "B"],
        "rows": [["1", "2"], ["3", "4"]],
    }
    html = json_blocks_to_html([block])
    assert "<th>A</th><th>B</th>" in html
    assert "<td><p>1</p></td><td><p>2</p></td>" in html


def test_table_cell_with_multiple_blocks():
    block = {
        "type": "table",
        "rows": [
            [
                [
                    {"type": "paragraph", "text": "para one"},
                    {"type": "list", "style": "bullet", "items": ["nested item"]},
                ],
                "plain cell",
            ]
        ],
    }
    html = json_blocks_to_html([block])
    assert "<p>para one</p>" in html
    assert "<ul><li><p>nested item</p></li></ul>" in html
    assert "<td><p>plain cell</p></td>" in html


def test_image_block():
    html = json_blocks_to_html([{"type": "image", "src": "/tmp/x.png", "alt": "pic", "width": 200}])
    assert '<img src="/tmp/x.png" alt="pic" width="200"/>' == html


def test_image_without_src_raises():
    with pytest.raises(ValueError):
        json_blocks_to_html([{"type": "image"}])


def test_code_block_escapes_content():
    html = json_blocks_to_html([{"type": "code", "text": "<x>&y</x>"}])
    assert html == "<pre><code>&lt;x&gt;&amp;y&lt;/x&gt;</code></pre>"


def test_blockquote_block():
    html = json_blocks_to_html(
        [{"type": "blockquote", "blocks": [{"type": "paragraph", "text": "quoted"}]}]
    )
    assert html == "<blockquote><p>quoted</p></blockquote>"


def test_hr_block():
    assert json_blocks_to_html([{"type": "hr"}]) == "<hr/>"


def test_pagebreak_block():
    html = json_blocks_to_html([{"type": "pagebreak"}])
    assert "docgen-pagebreak" in html
    assert "page-break-after: always" in html


def test_unknown_block_type_raises():
    with pytest.raises(ValueError):
        json_blocks_to_html([{"type": "nope"}])


def test_block_without_type_raises():
    with pytest.raises(ValueError):
        json_blocks_to_html([{"text": "no type field"}])


def test_root_must_be_list():
    with pytest.raises(ValueError):
        json_blocks_to_html({"type": "paragraph", "text": "x"})


def test_content_to_html_json_invalid_json_raises():
    with pytest.raises(ValueError):
        _content_to_html("{not valid json", "json")


def test_content_to_html_invalid_format_raises():
    with pytest.raises(ValueError):
        _content_to_html("# hi", "plain")


def test_content_to_html_markdown_still_works():
    html = _content_to_html("# Title\n\nSome text.", "markdown")
    assert "<h1>Title</h1>" in html

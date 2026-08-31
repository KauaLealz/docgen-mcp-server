import os

import pytest

from services import render_image_service


def _load_png_info(path):
    with open(path, "rb") as f:
        data = f.read()
    # PNG IHDR: width/height are bytes 16-24, big-endian; also check color
    # type byte (25) for alpha presence (6 = RGBA).
    assert data[:8] == b"\x89PNG\r\n\x1a\n"
    width = int.from_bytes(data[16:20], "big")
    height = int.from_bytes(data[20:24], "big")
    color_type = data[25]
    return width, height, color_type


def test_render_image_full_page_png(tmp_path):
    out = str(tmp_path / "img.png")
    render_image_service({
        "html": "<div style='width:100px;height:80px;background:#c00;'></div>",
        "css": "",
        "outputPath": out,
        "width": 100,
        "height": 80,
    })
    assert os.path.getsize(out) > 0
    w, h, _ = _load_png_info(out)
    assert (w, h) == (100, 80)


def test_render_image_with_selector_crops_to_element(tmp_path):
    out = str(tmp_path / "badge.png")
    render_image_service({
        "html": "<div class='badge' style='width:50px;height:50px;background:#0a0;'></div>",
        "css": "",
        "outputPath": out,
        "width": 300,
        "height": 300,
        "selector": ".badge",
    })
    w, h, _ = _load_png_info(out)
    assert (w, h) == (50, 50)


def test_render_image_missing_selector_raises(tmp_path):
    out = str(tmp_path / "x.png")
    with pytest.raises(ValueError):
        render_image_service({
            "html": "<div>no badge here</div>",
            "css": "",
            "outputPath": out,
            "width": 100,
            "height": 100,
            "selector": ".does-not-exist",
        })


def test_render_image_transparent_background_has_alpha_channel(tmp_path):
    out = str(tmp_path / "transparent.png")
    render_image_service({
        "html": "<div style='width:40px;height:40px;'></div>",
        "css": "",
        "outputPath": out,
        "width": 40,
        "height": 40,
        "transparentBackground": True,
    })
    _, _, color_type = _load_png_info(out)
    assert color_type == 6  # RGBA


def test_render_image_scale_increases_pixel_dimensions(tmp_path):
    out = str(tmp_path / "scaled.png")
    render_image_service({
        "html": "<div style='width:50px;height:50px;'></div>",
        "css": "",
        "outputPath": out,
        "width": 50,
        "height": 50,
        "scale": 2,
    })
    w, h, _ = _load_png_info(out)
    assert (w, h) == (100, 100)


def test_render_image_invalid_format_raises(tmp_path):
    out = str(tmp_path / "x.gif")
    with pytest.raises(ValueError):
        render_image_service({
            "html": "<div></div>", "css": "", "outputPath": out,
            "width": 10, "height": 10, "format": "gif",
        })


def test_render_image_jpeg_format(tmp_path):
    out = str(tmp_path / "x.jpg")
    render_image_service({
        "html": "<div style='width:20px;height:20px;background:#fff;'></div>",
        "css": "", "outputPath": out, "width": 20, "height": 20, "format": "jpeg",
    })
    with open(out, "rb") as f:
        header = f.read(3)
    assert header == b"\xff\xd8\xff"

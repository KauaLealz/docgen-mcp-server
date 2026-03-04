"""Generate charts via matplotlib and embed them in documents."""

import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from utils.file_utils import generate_output_path
from handlers.docx_handler import create_docx
from handlers.pdf_handler import create_pdf


def create_chart(
    chart_type: str,
    data: dict,
    title: str = "",
    xlabel: str = "",
    ylabel: str = "",
    width: float = 8.0,
    height: float = 5.0,
    output_path: str = "",
) -> str:
    """Generate a chart as a PNG image.

    Args:
        chart_type: One of "bar", "line", "pie", "horizontal_bar", "scatter", "area".
        data: Chart data. Structure depends on chart_type.
        title: Chart title.
        xlabel: X-axis label (ignored for pie).
        ylabel: Y-axis label (ignored for pie).
        width: Figure width in inches.
        height: Figure height in inches.
        output_path: Absolute path for the PNG output.

    Returns the absolute path of the generated PNG.
    """
    fig, ax = plt.subplots(figsize=(width, height))
    _render_chart(fig, ax, chart_type, data, title, xlabel, ylabel)

    dest = generate_output_path(output_path)
    fig.savefig(str(dest), dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return str(dest)


def create_chart_document(
    chart_type: str,
    data: dict,
    title: str = "",
    xlabel: str = "",
    ylabel: str = "",
    output_format: str = "docx",
    extra_sections: list[dict] | None = None,
    output_path: str = "",
) -> str:
    """Generate a chart and embed it in a docx or pdf document.

    The chart PNG is created as a temp file, embedded, then cleaned up.

    Returns the absolute path of the generated document.
    """
    fig, ax = plt.subplots(figsize=(8.0, 5.0))
    _render_chart(fig, ax, chart_type, data, title, xlabel, ylabel)

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp_path = tmp.name
    tmp.close()
    fig.savefig(tmp_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    try:
        sections: list[dict] = [
            {"type": "image", "path": tmp_path, "width_inches": 6.5, "caption": title},
        ]
        if extra_sections:
            sections.extend(extra_sections)

        doc_title = title or "Relatório com Gráfico"
        if output_format.lower() == "pdf":
            return create_pdf(doc_title, sections, output_path)
        return create_docx(doc_title, sections, output_path)
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _render_chart(fig, ax, chart_type, data, title, xlabel, ylabel):
    if chart_type == "bar":
        _render_bar(ax, data)
    elif chart_type == "horizontal_bar":
        _render_horizontal_bar(ax, data)
    elif chart_type == "line":
        _render_line(ax, data)
    elif chart_type == "pie":
        _render_pie(ax, data)
    elif chart_type == "scatter":
        _render_scatter(ax, data)
    elif chart_type == "area":
        _render_area(ax, data)
    else:
        raise ValueError(f"Tipo de gráfico não suportado: {chart_type}. "
                         f"Use: bar, horizontal_bar, line, pie, scatter, area")

    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", pad=12)
    if xlabel and chart_type != "pie":
        ax.set_xlabel(xlabel)
    if ylabel and chart_type != "pie":
        ax.set_ylabel(ylabel)

    if chart_type != "pie":
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    fig.tight_layout()


# ─── Chart Renderers ─────────────────────────────────────────────────────────

COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5",
          "#70AD47", "#264478", "#9B57A0", "#636363", "#FF5A5A"]


def _get_colors(n: int) -> list[str]:
    return [COLORS[i % len(COLORS)] for i in range(n)]


def _render_bar(ax, data: dict):
    labels = data.get("labels", [])
    datasets = data.get("datasets", [])
    if not datasets:
        return

    import numpy as np
    x = np.arange(len(labels))
    n = len(datasets)
    width = 0.7 / n

    for i, ds in enumerate(datasets):
        offset = (i - n / 2 + 0.5) * width
        values = ds.get("values", [])
        color = ds.get("color", COLORS[i % len(COLORS)])
        ax.bar(x + offset, values, width, label=ds.get("label", ""), color=color)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right" if len(labels) > 6 else "center")
    if n > 1:
        ax.legend()


def _render_horizontal_bar(ax, data: dict):
    labels = data.get("labels", [])
    datasets = data.get("datasets", [])
    if not datasets:
        return

    import numpy as np
    y = np.arange(len(labels))
    n = len(datasets)
    height = 0.7 / n

    for i, ds in enumerate(datasets):
        offset = (i - n / 2 + 0.5) * height
        values = ds.get("values", [])
        color = ds.get("color", COLORS[i % len(COLORS)])
        ax.barh(y + offset, values, height, label=ds.get("label", ""), color=color)

    ax.set_yticks(y)
    ax.set_yticklabels(labels)
    if n > 1:
        ax.legend()


def _render_line(ax, data: dict):
    labels = data.get("labels", [])
    datasets = data.get("datasets", [])

    for i, ds in enumerate(datasets):
        values = ds.get("values", [])
        color = ds.get("color", COLORS[i % len(COLORS)])
        marker = ds.get("marker", "o")
        ax.plot(labels, values, marker=marker, label=ds.get("label", ""),
                color=color, linewidth=2, markersize=5)

    if len(datasets) > 1:
        ax.legend()
    ax.grid(axis="y", alpha=0.3)
    if len(labels) > 6:
        plt.xticks(rotation=45, ha="right")


def _render_pie(ax, data: dict):
    labels = data.get("labels", [])
    values = data.get("values", [])
    colors = _get_colors(len(labels))
    explode = data.get("explode", [0] * len(labels))

    wedges, texts, autotexts = ax.pie(
        values, labels=labels, colors=colors, explode=explode,
        autopct="%1.1f%%", startangle=90, pctdistance=0.85,
    )
    for text in autotexts:
        text.set_fontsize(9)
    ax.axis("equal")


def _render_scatter(ax, data: dict):
    datasets = data.get("datasets", [])
    for i, ds in enumerate(datasets):
        x_vals = ds.get("x", [])
        y_vals = ds.get("y", ds.get("values", []))
        color = ds.get("color", COLORS[i % len(COLORS)])
        ax.scatter(x_vals, y_vals, label=ds.get("label", ""),
                   color=color, alpha=0.7, s=50)

    if len(datasets) > 1:
        ax.legend()
    ax.grid(alpha=0.3)


def _render_area(ax, data: dict):
    labels = data.get("labels", [])
    datasets = data.get("datasets", [])

    for i, ds in enumerate(datasets):
        values = ds.get("values", [])
        color = ds.get("color", COLORS[i % len(COLORS)])
        ax.fill_between(range(len(values)), values, alpha=0.3, color=color)
        ax.plot(range(len(values)), values, color=color,
                label=ds.get("label", ""), linewidth=1.5)

    if labels:
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0,
                           ha="right" if len(labels) > 6 else "center")
    if len(datasets) > 1:
        ax.legend()
    ax.grid(axis="y", alpha=0.3)

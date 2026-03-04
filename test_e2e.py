"""End-to-end test: generate docx/pdf/xlsx then read them back."""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault("DOCGEN_OUTPUT_DIR", os.path.join(os.path.dirname(__file__), "output"))
os.environ.setdefault("DOCGEN_ALLOWED_DIRS", "C:\\RTB;C:\\Users")

from handlers.docx_handler import create_docx, read_docx
from handlers.pdf_handler import create_pdf, read_pdf
from handlers.excel_handler import create_excel, read_excel
from handlers.markdown_handler import markdown_to_document
from handlers.chart_handler import create_chart, create_chart_document
from utils.file_utils import list_files

SAMPLE_SECTIONS = [
    {"type": "heading", "text": "1. Identificação", "level": 1},
    {"type": "paragraph", "text": "Este é um documento de teste gerado pelo MCP DocGen Server."},
    {"type": "table", "headers": ["Campo", "Valor"], "rows": [
        ["Serviço", "bwp-fixed-income-settlement"],
        ["Classe", "CotacaoTituloPublicoServiceImpl"],
        ["Tipo", "Hotfix"],
    ]},
    {"type": "heading", "text": "2. Código", "level": 1},
    {"type": "code_block", "code": "if (isCliente(origem)) {\n    this.atualizaOperacao(dto);\n}", "language": "java"},
    {"type": "heading", "text": "3. Itens", "level": 1},
    {"type": "list", "items": ["Item A", "Item B", "Item C"], "ordered": True},
    {"type": "paragraph", "text": "Fim do documento.", "bold": True},
]


def test_docx():
    print("=" * 60)
    print("TEST: create_docx + read_docx")
    print("=" * 60)
    path = create_docx("Teste DocGen DOCX", SAMPLE_SECTIONS)
    print(f"  Gerado: {path}")

    data = read_docx(path)
    print(f"  Paragrafos: {len(data['paragraphs'])}")
    print(f"  Tabelas: {len(data['tables'])}")
    print(f"  Imagens: {data['images_count']}")
    print(f"  Texto (primeiros 100 chars): {data['text'][:100]}...")
    if data["tables"]:
        t = data["tables"][0]
        print(f"  Tabela 0 headers: {t['headers']}")
        print(f"  Tabela 0 rows: {t['rows'][:2]}")
    print("  OK\n")
    return path


def test_pdf():
    print("=" * 60)
    print("TEST: create_pdf + read_pdf")
    print("=" * 60)
    path = create_pdf("Teste DocGen PDF", SAMPLE_SECTIONS)
    print(f"  Gerado: {path}")

    data = read_pdf(path)
    print(f"  Total paginas: {data['total_pages']}")
    for pg in data["pages"]:
        print(f"  Pagina {pg['page_number']}: {len(pg['text'])} chars, {len(pg['tables'])} tabelas")
    print("  OK\n")
    return path


def test_excel():
    print("=" * 60)
    print("TEST: create_excel + read_excel")
    print("=" * 60)
    sheets = [
        {
            "name": "Operações",
            "headers": ["ID", "Status", "Valor", "Data"],
            "rows": [
                ["OP-001", "PROCESSANDO", "15000.50", "2026-03-01"],
                ["OP-002", "REALIZADA", "23400.00", "2026-03-02"],
                ["OP-003", "APROVADO", "8750.25", "2026-03-03"],
            ],
        },
        {
            "name": "Resumo",
            "headers": ["Métrica", "Valor"],
            "rows": [
                ["Total operações", "3"],
                ["Volume total", "47150.75"],
            ],
        },
    ]
    path = create_excel("Teste DocGen Excel", sheets)
    print(f"  Gerado: {path}")

    data = read_excel(path)
    print(f"  Sheets: {data['sheet_names']}")
    for s in data["sheets"]:
        print(f"  Sheet '{s['name']}': {s['row_count']} rows, {s['column_count']} cols")
        print(f"    Headers: {s['headers']}")
        if s["rows"]:
            print(f"    First row: {s['rows'][0]}")
    print("  OK\n")
    return path


def test_markdown_to_docx():
    print("=" * 60)
    print("TEST: markdown_to_document (docx)")
    print("=" * 60)
    md = """# Relatório de Hotfix

## Contexto

Este documento foi gerado a partir de **Markdown** pelo MCP DocGen Server.

## Impacto

| Métrica | Valor |
|---------|-------|
| Operações afetadas | 5 |
| Serviço | bwp-fixed-income-settlement |

## Código corrigido

```java
if (isCliente(origem)) {
    this.atualizaOperacao(dto);
}
```

## Ações

1. Equalizar branch com develop
2. Validar em UAT
3. Deploy em produção

- Item sem número A
- Item sem número B

---

**Documento finalizado.**
"""
    path = markdown_to_document(md, output_format="docx")
    print(f"  Gerado DOCX: {path}")

    data = read_docx(path)
    print(f"  Paragrafos: {len(data['paragraphs'])}")
    print(f"  Tabelas: {len(data['tables'])}")
    print("  OK\n")

    path_pdf = markdown_to_document(md, output_format="pdf", title="Relatório Hotfix PDF")
    print(f"  Gerado PDF: {path_pdf}")
    print("  OK\n")


def test_chart_png():
    print("=" * 60)
    print("TEST: create_chart (bar + line + pie)")
    print("=" * 60)

    bar_path = create_chart(
        chart_type="bar",
        data={
            "labels": ["Jan", "Fev", "Mar", "Abr", "Mai"],
            "datasets": [
                {"label": "Operações", "values": [120, 150, 180, 90, 200]},
                {"label": "Resgates", "values": [30, 45, 60, 20, 55]},
            ],
        },
        title="Operações vs Resgates",
        xlabel="Mês",
        ylabel="Quantidade",
    )
    print(f"  Bar chart: {bar_path}")

    line_path = create_chart(
        chart_type="line",
        data={
            "labels": ["Seg", "Ter", "Qua", "Qui", "Sex"],
            "datasets": [
                {"label": "Latência (ms)", "values": [120, 95, 110, 88, 72]},
            ],
        },
        title="Latência Semanal",
        ylabel="ms",
    )
    print(f"  Line chart: {line_path}")

    pie_path = create_chart(
        chart_type="pie",
        data={
            "labels": ["LTN", "NTN-F", "NTN-B", "LFT"],
            "values": [35, 25, 30, 10],
        },
        title="Distribuição por Tipo de Título",
    )
    print(f"  Pie chart: {pie_path}")
    print("  OK\n")


def test_chart_document():
    print("=" * 60)
    print("TEST: create_chart_document (docx + pdf)")
    print("=" * 60)

    docx_path = create_chart_document(
        chart_type="bar",
        data={
            "labels": ["Portal", "FTS"],
            "datasets": [{"label": "Total", "values": [735, 735]}],
        },
        title="Batimento Portal x FTS",
        ylabel="Operações",
        output_format="docx",
        extra_sections=[
            {"type": "paragraph", "text": "O batimento Portal (735) = FTS (735) confirma zero divergências."},
        ],
    )
    print(f"  Chart DOCX: {docx_path}")

    pdf_path = create_chart_document(
        chart_type="line",
        data={
            "labels": ["D-5", "D-4", "D-3", "D-2", "D-1", "Hoje"],
            "datasets": [
                {"label": "Divergências", "values": [5, 3, 0, 0, 0, 0]},
            ],
        },
        title="Evolução de Divergências Pós-Hotfix",
        ylabel="Quantidade",
        output_format="pdf",
    )
    print(f"  Chart PDF: {pdf_path}")
    print("  OK\n")


def test_list_files():
    print("=" * 60)
    print("TEST: list_generated_files")
    print("=" * 60)
    files = list_files()
    for f in files:
        print(f"  {f['name']} ({f['size_human']}) - {f['modified']}")
    print(f"  Total: {len(files)} arquivos")
    print("  OK\n")


if __name__ == "__main__":
    print("\n MCP DocGen Server - E2E Tests\n")
    try:
        test_docx()
        test_pdf()
        test_excel()
        test_markdown_to_docx()
        test_chart_png()
        test_chart_document()
        test_list_files()
        print("=" * 60)
        print("TODOS OS TESTES PASSARAM")
        print("=" * 60)
    except Exception as e:
        print(f"\nFALHA: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

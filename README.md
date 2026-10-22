# docgen-mcp-server

> Servidor [MCP](https://modelcontextprotocol.io) para **ler, escrever, editar e renderizar documentos** — DOCX, PDF, PPTX, XLSX, CSV, ODT — além de HTML/CSS → PDF/slides/imagem, diff, busca e empacotamento.

## O que faz

Dá a um agente um conjunto de ferramentas de escritório sobre arquivos locais: extrair conteúdo como Markdown, gerar documentos e planilhas, renderizar HTML/CSS em PDF (Chromium), slides PowerPoint ou imagens, e operações utilitárias (patch, diff, busca, zip).

| Tool | O que faz |
|------|-----------|
| `read_doc` | Lê `.docx`/`.pdf`/`.odt`/`.pptx` e devolve Markdown. |
| `read_sheet` | Lê planilha (`.xlsx`/`.csv`) como Markdown ou JSON (`range`, `sheetName`). |
| `read_archive` | Lista a árvore de um `.zip` (com filtro opcional). |
| `write_doc` | Gera `.docx` ou `.pdf` a partir de Markdown ou blocos JSON; suporta template `{{campo}}`. |
| `write_sheet` | Gera `.xlsx` ou `.csv` a partir de dados. |
| `render_slide` | HTML/CSS → apresentação `.pptx` (um slide por `<section>`). |
| `render_page` | HTML/CSS → PDF A4 (Chromium). |
| `render_image` | HTML/CSS → imagem PNG/JPEG. |
| `patch_doc` | Edita PDF (merge/split/watermark) e DOCX (replace_text). |
| `patch_sheet` | Atualiza células de um `.xlsx`. |
| `scan_dir` | Busca por regex em diretório/arquivos/ZIPs. |
| `diff_file` | Diff de texto ou de planilhas. |
| `bundle_zip` | Compacta uma lista de arquivos num ZIP. |

## Requisitos

- **Python 3.11+**.
- **Chromium** — necessário só para renderização a partir de HTML (`render_page`, `render_image` e `write_doc` em PDF). Instale uma vez com `playwright install chromium`. As tools de DOCX/XLSX/PPTX/leitura **não** precisam de navegador.

## Instalação

```bash
# Executar direto, via uv:
uvx docgen-mcp-server

# Ou instalar no sistema:
uv tool install docgen-mcp-server
# (alternativa: pip install docgen-mcp-server)

# Uma vez, se for usar renderização HTML (PDF/imagem):
playwright install chromium
```

## Configuração

Aponte o cliente MCP para o servidor. Ele **roda sem nenhuma configuração obrigatória** — não precisa de chave nem credencial:

```json
{
  "mcpServers": {
    "docgen": {
      "command": "uvx",
      "args": ["docgen-mcp-server"],
      "env": {}
    }
  }
}
```

O mesmo bloco vale para Cursor, Claude Desktop, Claude Code e outras IDEs com MCP por stdio.

### Variáveis de ambiente (todas opcionais)

| Variável | Efeito | Default |
|----------|--------|---------|
| `DOCGEN_ALLOWED_ROOTS` | Pastas (separadas por vírgula) às quais limitar toda leitura/escrita. Fora delas → erro. | sem restrição |
| `DOCGEN_SCAN_MAX_MATCHES` | Teto de correspondências em `scan_dir` (1–50000). | 500 |
| `DOCGEN_READ_SHEET_MAX_ROWS` | Teto de linhas em `read_sheet` sem `maxRows` (1–100000). | 10000 |

## Segurança

- **Path traversal bloqueado** (rejeita `..`) e **limite de 50 MB** por arquivo lido.
- **Allowlist opcional** de raízes via `DOCGEN_ALLOWED_ROOTS`.
- **Escrita bloqueada** em diretórios protegidos do sistema (`C:\Windows`, `Program Files`, `/etc`, `/usr`…) e em pastas sensíveis do home (`.ssh`, `.aws`, `.gnupg`, `.kube`, `.docker`…).

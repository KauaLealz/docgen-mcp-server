# DocGen MCP Server

## O que é

Servidor [Model Context Protocol (MCP)](https://modelcontextprotocol.io) via **stdio** para leitura e escrita de documentos, planilhas, renderização HTML→PDF (Playwright) e utilitários de arquivo. É distribuído no PyPI como **`docgen-mcp-server`** e requer **Python 3.11+**.

Recomenda-se executar com **`uvx docgen-mcp-server`**. Para travar em uma versão, use `uvx --from docgen-mcp-server==0.2.2 docgen-mcp-server`.

Na primeira execução, o Playwright pode precisar de um navegador Chromium/Chrome instalado (tools `render_page` e `render_image`). `render_slide` gera o arquivo PowerPoint diretamente, sem navegador.

## Ferramentas

| Módulo | Ferramenta | Descrição resumida |
|--------|------------|-------------------|
| **read_** | `read_doc` | `.docx`/`.pdf`/`.odt`/`.pptx` → Markdown; `previewOnly` / `maxChars` limitam saída. |
| | `read_sheet` | Planilhas → JSON ou Markdown; `range` tipo `A1:D10`; `previewOnly` / `maxRows`. |
| | `read_archive` | Lista árvore de entradas em `.zip`. |
| **write_** | `write_doc` | `.docx`/`.pdf`; **Markdown** (`#`, listas, \`\`\`) ou blocos JSON tipados; template `{{campo}}`. |
| | `write_sheet` | `.xlsx` ou `.csv`; `append` em `.xlsx` para logs. |
| **render_** | `render_slide` | HTML/CSS → PowerPoint `.pptx` (use `<section>` ou `.slide` por slide; `assetBasePath` resolve imagens relativas e `footerHtml` aceita `{slideNumber}`/`{slideCount}`). |
| | `render_page` | HTML/CSS → PDF A4 (índice opcional). |
| **patch_** | `patch_doc` | PDF: merge, split, watermark; DOCX: `replace_text` em XML. |
| | `patch_sheet` | Atualiza células em `.xlsx`. |
| **system_** | `scan_dir` | Busca por regex em diretório ou em arquivos/ZIPs. |
| | `diff_file` | Diff texto ou dados (planilhas). |
| | `bundle_zip` | Compacta lista de arquivos em um ZIP. |

### Segurança

- Sem `..` nos caminhos; leitura limitada por tamanho de ficheiro.
- **Escrita bloqueada** em pastas do sistema (ex.: `Windows`, `Program Files`, `.ssh`, `.aws` no home).
- Opcional: **`DOCGEN_ALLOWED_ROOTS`** — lista separada por vírgulas de pastas absolutas; só é permitido ler/escrever dentro delas (útil em monorepos/CI).

### Variáveis de ambiente (opcional)

| Variável | Efeito |
|----------|--------|
| `DOCGEN_ALLOWED_ROOTS` | Ex.: `C:\repo\my-app,C:\tmp` — restrição de caminhos. |
| `DOCGEN_SCAN_MAX_MATCHES` | Máximo de correspondências em `scan_dir` (padrão 500). |
| `DOCGEN_READ_SHEET_MAX_ROWS` | Teto de linhas de dados em `read_sheet` quando não usas `maxRows` (padrão 10000). |

Erros das tools devolvem **`structuredContent`** com `ok: false`, `code`, `tool`, `message` e às vezes `hint`.

## Como usar nos clientes (recomendado: PyPI)

Em qualquer cliente MCP com transporte **stdio**:

- **Comando**: `uvx`.
- **Argumentos**: `["docgen-mcp-server"]` (ou `["--from", "docgen-mcp-server==0.2.2", "docgen-mcp-server"]` para uma versão fixa).
- **Variáveis de ambiente**: opcionais; veja as variáveis do Playwright/Chromium se precisar de proxy ou caminho de browser.

Exemplo (Cursor, VS Code com MCP, Claude Desktop, etc.):

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

No repositório há um exemplo em [`.cursor/mcp.json.example`](.cursor/mcp.json.example).

### Cursor

**Configurações → MCP** (ou JSON de MCP do projeto): use `command`, `args` e `env` como acima.

### Claude Desktop

Mesmo esquema de `command`, `args` e `env`. Detalhes de caminho do arquivo de configuração variam por SO; veja a [documentação da Anthropic sobre MCP](https://docs.anthropic.com/en/docs/agents-and-tools/mcp).

### Problemas comuns na instalação

Se `uvx` não for encontrado, instale o [uv](https://docs.astral.sh/uv/getting-started/installation/). Para atualizar uma instalação em cache, use `uvx --refresh docgen-mcp-server`.

Se a ferramenta de renderização reclamar de navegador, instale o Chromium com `playwright install chromium` ou configure o caminho de um Chrome já instalado.

## Desenvolvimento a partir do clone

```bash
git clone <repo>
cd docgen-mcp-server
uv sync
uv run docgen-mcp-server
```

Sem build prévio (local):

```bash
uv run docgen-mcp-server
```

| Script | Ação |
|--------|------|
| `uv sync` | Instala dependências do projeto |
| `uv run docgen-mcp-server` | Executa o servidor |
| `uv build` | Gera os artefatos para publicação |

## Publicação no PyPI (mantenedores)

```bash
uv build
uv publish --dry-run
uv publish
```

## Licença

ISC

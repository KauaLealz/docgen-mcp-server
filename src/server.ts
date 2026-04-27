import "./stdio-guard.js"
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { registerAllTools } from "./tools/index.js"

export function createMcpServer(): McpServer {
  const server = new McpServer(
    { name: "docgen", version: "3.0.0" },
    {
      instructions:
        "Docgen MCP: ingestão (read_doc, read_sheet, read_archive), geração (write_doc, write_sheet), renderização HTML (render_slide, render_page), edição (patch_doc, patch_sheet), sistema (scan_dir, diff_file, bundle_zip). Caminhos devem ser absolutos quando o cliente gravar arquivos. Renderização usa Puppeteer/Chromium.",
    },
  )
  registerAllTools(server)
  return server
}

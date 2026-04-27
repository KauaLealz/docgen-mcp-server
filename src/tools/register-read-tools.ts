import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { safeTool } from "../shared/tool-runner.js"
import { readArchive } from "../services/read-archive.js"
import { readDoc } from "../services/read-doc.js"
import { readSheet } from "../services/read-sheet.js"

export function registerReadTools(server: McpServer): void {
  server.registerTool(
    "read_doc",
    {
      description:
        "Extrai texto de .docx, .pdf ou .odt. Markdown; extractImages (DOCX). previewOnly/maxChars limitam tamanho da saída para o LLM.",
      inputSchema: {
        path: z.string().describe("Caminho absoluto do arquivo"),
        includeComments: z.boolean().optional(),
        extractImages: z.boolean().optional(),
        previewOnly: z
          .boolean()
          .optional()
          .describe("Se true, limita texto (~12k chars); combine com maxChars."),
        maxChars: z.coerce.number().int().positive().optional().describe("Teto de caracteres no texto extraído"),
      },
    },
    async (args: {
      path: string
      includeComments?: boolean
      extractImages?: boolean
      previewOnly?: boolean
      maxChars?: number
    }) =>
      safeTool("read_doc", async () => {
        const r = await readDoc(args.path, {
          includeComments: args.includeComments,
          extractImages: args.extractImages,
          maxChars: args.maxChars,
          previewOnly: args.previewOnly,
        })
        const text = r.markdown + (r.note ? `\n\n_${r.note}_` : "")
        return {
          text,
          structured: {
            markdown: r.markdown,
            truncated: r.truncated ?? false,
            imagePaths: args.extractImages ? r.imagePaths : undefined,
            tool: "read_doc",
          },
        }
      }),
  )

  server.registerTool(
    "read_sheet",
    {
      description:
        "Lê .xlsx, .csv ou .ods. asJson: array de objetos ou tabela Markdown. range ex.: A1:D10. previewOnly/maxRows truncam linhas de dados.",
      inputSchema: {
        path: z.string(),
        sheetName: z.string().optional(),
        range: z.string().optional().describe("Intervalo Excel, ex.: A1:D10"),
        asJson: z.boolean(),
        previewOnly: z.boolean().optional().describe("Trunca a ~500 linhas de dados"),
        maxRows: z.coerce.number().int().positive().optional().describe("Máximo de linhas de dados (excl. cabeçalho)"),
      },
    },
    async (args: {
      path: string
      sheetName?: string
      range?: string
      asJson: boolean
      previewOnly?: boolean
      maxRows?: number
    }) =>
      safeTool("read_sheet", async () => {
        const out = await readSheet(args.path, {
          sheetName: args.sheetName,
          range: args.range,
          asJson: args.asJson,
          previewOnly: args.previewOnly,
          maxRows: args.maxRows,
        })
        const text = out.asJson
          ? JSON.stringify(out.jsonRows ?? [], null, 2)
          : (out.markdown ?? "")
        return {
          text,
          structured: {
            truncated: out.truncated,
            rowLimit: out.rowLimit,
            tool: "read_sheet",
          },
        }
      }),
  )

  server.registerTool(
    "read_archive",
    {
      description:
        "Lista estrutura de um arquivo .zip (árvore de caminhos). Opcionalmente filtre entries com pattern tipo glob simples (*).",
      inputSchema: {
        path: z.string(),
        pattern: z.string().optional(),
      },
    },
    async (args: { path: string; pattern?: string }) =>
      safeTool("read_archive", async () => await readArchive(args.path, args.pattern)),
  )
}

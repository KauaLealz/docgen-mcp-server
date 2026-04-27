import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { safeTool } from "../shared/tool-runner.js"
import { writeDoc } from "../services/write-doc.js"
import { writeSheet } from "../services/write-sheet.js"

export function registerWriteTools(server: McpServer): void {
  server.registerTool(
    "write_doc",
    {
      description:
        "Cria .docx ou PDF. contentFormat markdown (padrão): títulos #–######, listas -/*, blocos ```. plain: legado. Template DOCX com {{campo}}.",
      inputSchema: {
        path: z.string().describe("Caminho absoluto de saída"),
        type: z.enum(["docx", "pdf"]),
        content: z.string(),
        contentFormat: z.enum(["markdown", "plain"]).optional().describe("markdown (default) ou plain"),
        templatePath: z.string().optional(),
        mergeFields: z.record(z.unknown()).optional(),
      },
    },
    async (args: {
      path: string
      type: "docx" | "pdf"
      content: string
      contentFormat?: "markdown" | "plain"
      templatePath?: string
      mergeFields?: Record<string, unknown>
    }) =>
      safeTool("write_doc", async () => {
        return await writeDoc({
          path: args.path,
          type: args.type,
          content: args.content,
          contentFormat: args.contentFormat ?? "markdown",
          templatePath: args.templatePath,
          mergeFields: args.mergeFields,
        })
      }),
  )

  server.registerTool(
    "write_sheet",
    {
      description:
        "Cria .xlsx ou .csv (extensão no path). data: objetos ou linhas. append: só .xlsx — acrescenta linhas sem novo cabeçalho.",
      inputSchema: {
        path: z.string(),
        data: z.array(z.any()),
        columns: z.record(z.string()).optional(),
        freezePanes: z.boolean().optional(),
        append: z.boolean().optional().describe("Anexar linhas a .xlsx existente"),
      },
    },
    async (args: {
      path: string
      data: unknown[]
      columns?: Record<string, string>
      freezePanes?: boolean
      append?: boolean
    }) =>
      safeTool("write_sheet", async () => {
        return await writeSheet({
          path: args.path,
          data: args.data,
          columns: args.columns,
          freezePanes: args.freezePanes,
          append: args.append,
        })
      }),
  )
}

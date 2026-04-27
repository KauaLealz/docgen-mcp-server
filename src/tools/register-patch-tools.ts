import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { safeTool } from "../shared/tool-runner.js"
import { patchDoc } from "../services/patch-doc.js"
import { patchSheet } from "../services/patch-sheet.js"

export function registerPatchTools(server: McpServer): void {
  server.registerTool(
    "patch_doc",
    {
      description:
        "merge: path=saida.pdf, payload.sources=[pdfs]. split|watermark|replace_text: path=arquivo existente. replace_text DOCX substitui texto em word/document.xml.",
      inputSchema: {
        path: z.string(),
        action: z.enum(["merge", "split", "watermark", "replace_text"]),
        payload: z.any().describe("merge: { sources: string[] }; split: { outputDir?: string }; watermark: { text, opacity? }; replace_text: { replacements:[{from,to}] }"),
      },
    },
    async (args: { path: string; action: "merge" | "split" | "watermark" | "replace_text"; payload: unknown }) =>
      safeTool("patch_doc", async () => {
        return await patchDoc({
          path: args.path,
          action: args.action,
          payload: args.payload,
        })
      }),
  )

  server.registerTool(
    "patch_sheet",
    {
      description:
        "Atualiza células em .xlsx existente (usa primeira planilha). updates usa endereços tipo A1, B2.",
      inputSchema: {
        path: z.string(),
        updates: z.array(
          z.object({
            cell: z.string(),
            value: z.unknown(),
            style: z.record(z.unknown()).optional(),
          }),
        ),
      },
    },
    async (args) =>
      safeTool("patch_sheet", async () => {
        const { path: filePath, updates } = args as {
          path: string
          updates: { cell: string; value: unknown; style?: Record<string, unknown> }[]
        }
        return await patchSheet(filePath, updates)
      }),
  )
}

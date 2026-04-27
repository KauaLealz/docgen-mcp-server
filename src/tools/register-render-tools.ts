import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { safeTool } from "../shared/tool-runner.js"
import { renderPage, renderSlide } from "../services/render-html.js"

export function registerRenderTools(server: McpServer): void {
  server.registerTool(
    "render_slide",
    {
      description:
        "HTML/CSS para slides. Use divs .slide por página. Saída PDF ou ZIP com um HTML por slide. Requer outputPath absoluto (.pdf ou .zip).",
      inputSchema: {
        html: z.string(),
        css: z.string(),
        format: z.enum(["pdf", "zip"]),
        aspectRatio: z.enum(["16:9", "4:3"]),
        outputPath: z.string().describe("Caminho absoluto do PDF ou ZIP de saída"),
      },
    },
    async (args: {
      html: string
      css: string
      format: "pdf" | "zip"
      aspectRatio: "16:9" | "4:3"
      outputPath: string
    }) =>
      safeTool("render_slide", async () => {
        return await renderSlide({
          html: args.html,
          css: args.css,
          format: args.format,
          aspectRatio: args.aspectRatio,
          outputPath: args.outputPath,
        })
      }),
  )

  server.registerTool(
    "render_page",
    {
      description:
        "HTML/CSS para documento tipo relatório em PDF (A4). generateTOC opcional. outputPath absoluto .pdf.",
      inputSchema: {
        html: z.string(),
        css: z.string(),
        outputPath: z.string(),
        generateTOC: z.boolean().optional(),
        margins: z
          .object({
            top: z.string().optional(),
            right: z.string().optional(),
            bottom: z.string().optional(),
            left: z.string().optional(),
          })
          .optional(),
      },
    },
    async (args: {
      html: string
      css: string
      outputPath: string
      generateTOC?: boolean
      margins?: { top?: string; right?: string; bottom?: string; left?: string }
    }) =>
      safeTool("render_page", async () => {
        return await renderPage({
          html: args.html,
          css: args.css,
          outputPath: args.outputPath,
          generateTOC: args.generateTOC,
          margins: args.margins,
        })
      }),
  )
}

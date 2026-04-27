import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { getScanMaxMatches } from "../config/limits.js"
import { safeTool } from "../shared/tool-runner.js"
import { bundleZip } from "../services/bundle-zip.js"
import { diffFile } from "../services/diff-file.js"
import { scanDir } from "../services/scan-dir.js"

export function registerSystemTools(server: McpServer): void {
  server.registerTool(
    "scan_dir",
    {
      description:
        "Busca regex em diretório ou arquivo / ZIP. Limite de resultados: maxMatches ou DOCGEN_SCAN_MAX_MATCHES (padrão 500).",
      inputSchema: {
        path: z.string(),
        query: z.string().describe("Expressão regular (JavaScript)"),
        recursive: z.boolean(),
        maxMatches: z.coerce.number().int().positive().optional(),
      },
    },
    async (args: { path: string; query: string; recursive: boolean; maxMatches?: number }) =>
      safeTool("scan_dir", async () => {
        const cap = args.maxMatches ?? getScanMaxMatches()
        const r = await scanDir(args.path, args.query, args.recursive, cap)
        const lines = r.matches
        const text =
          lines.join("\n") +
          (r.truncated ? `\n\n_[Lista truncada: máximo ${cap} correspondências]_` : "")
        return {
          text: text || "Nenhuma correspondência.",
          structured: {
            count: r.matches.length,
            truncated: r.truncated,
            maxMatches: cap,
            matches: lines,
            tool: "scan_dir",
          },
        }
      }),
  )

  server.registerTool(
    "diff_file",
    {
      description:
        "Compara dois arquivos. mode=text (UTF-8 linha a linha); mode=data para planilhas (.xlsx/.csv/.ods) como estrutura JSON.",
      inputSchema: {
        pathA: z.string(),
        pathB: z.string(),
        mode: z.enum(["text", "data"]),
      },
    },
    async (args: { pathA: string; pathB: string; mode: "text" | "data" }) =>
      safeTool("diff_file", async () => diffFile(args.pathA, args.pathB, args.mode)),
  )

  server.registerTool(
    "bundle_zip",
    {
      description: "Compacta vários arquivos em um .zip. outputName pode ser caminho absoluto ou nome no diretório atual.",
      inputSchema: {
        files: z.array(z.string()),
        outputName: z.string(),
      },
    },
    async (args: { files: string[]; outputName: string }) =>
      safeTool("bundle_zip", async () => bundleZip(args.files, args.outputName)),
  )
}

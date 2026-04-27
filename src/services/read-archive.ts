import fs from "node:fs"
import path from "node:path"
import yauzl from "../shared/yauzl-cjs.js"
import { validatePath } from "../shared/security.js"

function globToRegex(pattern: string): RegExp {
  const esc = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*\*/g, "{{GLOBSTAR}}").replace(/\*/g, "[^/]*").replace(/\?/g, ".").replace(/\{\{GLOBSTAR\}\}/g, ".*")
  return new RegExp(`^${esc}$`, "i")
}

function buildTree(entries: string[]): string {
  const root: Record<string, unknown> = {}
  for (const e of entries.sort()) {
    const parts = e.split("/").filter(Boolean)
    let cur = root
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i]!
      if (i === parts.length - 1) {
        cur[p] = "(file)"
      } else {
        if (!cur[p] || typeof cur[p] !== "object") cur[p] = {}
        cur = cur[p] as Record<string, unknown>
      }
    }
  }

  const lines: string[] = []
  function walk(obj: Record<string, unknown>, prefix = ""): void {
    const keys = Object.keys(obj).sort()
    for (const k of keys) {
      const v = obj[k]
      const isFile = typeof v === "string" && v === "(file)"
      const line = `${prefix}${k}`
      lines.push(line + (isFile ? "" : "/"))
      if (!isFile && v && typeof v === "object") {
        walk(v as Record<string, unknown>, `${prefix}${k}/`)
      }
    }
  }
  walk(root)
  return lines.join("\n") || "(vazio)"
}

export async function readArchive(filePath: string, pattern?: string): Promise<string> {
  const resolved = validatePath(filePath, true)
  const ext = path.extname(resolved).toLowerCase()
  if (ext !== ".zip") {
    throw new Error(`read_archive suporta apenas .zip nesta versão. Recebido: ${ext}`)
  }

  const re = pattern?.trim() ? globToRegex(pattern.trim()) : null

  const entries = await new Promise<string[]>((resolve, reject) => {
    const names: string[] = []
    yauzl.open(resolved, { lazyEntries: true }, (err, zipfile) => {
      if (err || !zipfile) {
        reject(err ?? new Error("Falha ao abrir ZIP"))
        return
      }
      zipfile.readEntry()
      zipfile.on("entry", (entry) => {
        const name = entry.fileName
        if (!re || re.test(name.replace(/\\/g, "/"))) {
          names.push(name.replace(/\\/g, "/"))
        }
        zipfile.readEntry()
      })
      zipfile.on("end", () => resolve(names))
      zipfile.on("error", reject)
    })
  })

  const tree = buildTree(entries)
  const header = `Arquivo: ${resolved}\nEntradas${re ? ` (filtro: ${pattern})` : ""}: ${entries.length}\n\n`
  return header + tree
}

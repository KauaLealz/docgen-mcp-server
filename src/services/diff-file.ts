import fs from "node:fs"
import path from "node:path"
import { diffLines } from "diff"
import * as XLSX from "xlsx"
import { validatePath } from "../shared/security.js"

export function diffFile(
  pathA: string,
  pathB: string,
  mode: "text" | "data",
): string {
  const a = validatePath(pathA, true)
  const b = validatePath(pathB, true)

  if (mode === "data") {
    const extA = path.extname(a).toLowerCase()
    const extB = path.extname(b).toLowerCase()
    if (extA !== extB) {
      return "diff data: as extensões devem coincidir para modo data."
    }
    if (extA === ".xlsx" || extA === ".xls" || extA === ".csv" || extA === ".ods") {
      const wa = XLSX.read(fs.readFileSync(a), { type: "buffer" })
      const wb = XLSX.read(fs.readFileSync(b), { type: "buffer" })
      const ja = JSON.stringify(wa.SheetNames.reduce<Record<string, unknown>>((acc, n) => {
        acc[n] = XLSX.utils.sheet_to_json(wa.Sheets[n]!, { header: 1 })
        return acc
      }, {}))
      const jb = JSON.stringify(wb.SheetNames.reduce<Record<string, unknown>>((acc, n) => {
        acc[n] = XLSX.utils.sheet_to_json(wb.Sheets[n]!, { header: 1 })
        return acc
      }, {}))
      if (ja === jb) return "Planilhas equivalentes (modo data)."
      const parts = diffLines(ja, jb)
      const out: string[] = []
      for (const p of parts) {
        const prefix = p.added ? "+" : p.removed ? "-" : " "
        for (const line of p.value.split(/\n/).filter(Boolean)) {
          out.push(`${prefix} ${line}`)
        }
      }
      return out.slice(0, 5000).join("\n") + (out.length > 5000 ? "\n... (truncado)" : "")
    }
    return "Modo data: apenas planilhas tabulares (.xlsx, .csv, .ods) suportadas."
  }

  let fa: string
  let fb: string
  try {
    fa = fs.readFileSync(a, "utf8")
    fb = fs.readFileSync(b, "utf8")
  } catch {
    return "Modo text: um dos arquivos não é texto UTF-8 legível; use modo data para binários tabulares ou converta antes."
  }
  const parts = diffLines(fa, fb)
  const out: string[] = []
  for (const p of parts) {
    const prefix = p.added ? "+" : p.removed ? "-" : " "
    for (const line of p.value.split(/\n/)) {
      out.push(`${prefix} ${line}`)
    }
  }
  return out.slice(0, 8000).join("\n") + (out.length > 8000 ? "\n... (truncado)" : "")
}

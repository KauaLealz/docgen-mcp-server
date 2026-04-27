import fs from "node:fs"
import path from "node:path"
import * as XLSX from "xlsx"
import { getReadSheetMaxRows } from "../config/limits.js"
import { validatePath } from "../shared/security.js"

function stripBom(buf: Buffer): Buffer {
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return Buffer.from(buf.subarray(3))
  }
  return Buffer.from(buf)
}

function sheetToMarkdown(rows: unknown[][]): string {
  if (rows.length === 0) return ""
  const esc = (c: unknown) =>
    String(c ?? "")
      .replace(/\|/g, "\\|")
      .replace(/\n/g, " ")
  const header = rows[0]!.map(esc)
  const sep = header.map(() => "---")
  const lines = [
    `| ${header.join(" | ")} |`,
    `| ${sep.join(" | ")} |`,
    ...rows.slice(1).map((r) => `| ${r.map(esc).join(" | ")} |`),
  ]
  return lines.join("\n")
}

function parseRange(
  rangeStr: string | undefined,
  maxRow: number,
  maxCol: number,
): { r0: number; r1: number; c0: number; c1: number } | null {
  if (!rangeStr?.trim()) return null
  const m = rangeStr.trim().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i)
  if (!m) return null
  const colToIdx = (letters: string) => {
    let n = 0
    for (const ch of letters.toUpperCase()) {
      n = n * 26 + (ch.charCodeAt(0) - 64)
    }
    return n - 1
  }
  const c0 = colToIdx(m[1]!)
  const r0 = parseInt(m[2]!, 10) - 1
  const c1 = colToIdx(m[3]!)
  const r1 = parseInt(m[4]!, 10) - 1
  return {
    r0: Math.max(0, r0),
    r1: Math.min(maxRow - 1, r1),
    c0: Math.max(0, c0),
    c1: Math.min(maxCol - 1, c1),
  }
}

export type ReadSheetOutput = {
  asJson: boolean
  markdown?: string
  jsonRows?: Record<string, unknown>[]
  truncated: boolean
  rowLimit: number
}

export async function readSheet(
  filePath: string,
  options: {
    sheetName?: string
    range?: string
    asJson: boolean
    maxRows?: number
    previewOnly?: boolean
  },
): Promise<ReadSheetOutput> {
  const resolved = validatePath(filePath, true)
  const ext = path.extname(resolved).toLowerCase()

  const fileBuf = fs.readFileSync(resolved)
  const raw: Buffer = ext === ".csv" ? stripBom(Buffer.from(fileBuf)) : Buffer.from(fileBuf)

  let workbook: XLSX.WorkBook
  if (ext === ".csv") {
    workbook = XLSX.read(raw, { type: "buffer", raw: true, codepage: 65001 })
  } else if (ext === ".ods") {
    workbook = XLSX.read(raw, { type: "buffer", cellDates: true })
  } else if (ext === ".xlsx" || ext === ".xlsm" || ext === ".xls") {
    workbook = XLSX.read(raw, { type: "buffer", cellDates: true })
  } else {
    throw new Error(`Extensão não suportada para read_sheet: ${ext}. Use .xlsx, .csv ou .ods.`)
  }

  const sheetNames = workbook.SheetNames
  if (sheetNames.length === 0) throw new Error("Planilha vazia ou sem folhas.")

  const name =
    options.sheetName && workbook.Sheets[options.sheetName]
      ? options.sheetName
      : sheetNames[0]!

  const sheet = workbook.Sheets[name]
  if (!sheet) throw new Error(`Folha não encontrada: ${options.sheetName ?? "(default)"}`)

  const ref = sheet["!ref"]
  const envCap = getReadSheetMaxRows()
  const cap =
    options.maxRows != null && options.maxRows > 0
      ? Math.min(options.maxRows, envCap)
      : options.previewOnly
        ? Math.min(500, envCap)
        : envCap

  if (!ref) {
    return {
      asJson: options.asJson,
      markdown: "",
      jsonRows: options.asJson ? [] : undefined,
      truncated: false,
      rowLimit: cap,
    }
  }

  const decoded = XLSX.utils.decode_range(ref)
  if (options.range?.trim()) {
    const test = parseRange(options.range, decoded.e.r + 100, decoded.e.c + 100)
    if (!test) {
      throw new Error(
        `range inválido: "${options.range}". Use o formato Excel com colunas e linhas, ex.: A1:D10 ou B2:B50.`,
      )
    }
  }

  const bounds = parseRange(options.range, decoded.e.r + 1, decoded.e.c + 1)
  let r0 = decoded.s.r
  let r1 = decoded.e.r
  let c0 = decoded.s.c
  let c1 = decoded.e.c
  if (bounds) {
    r0 = bounds.r0
    r1 = bounds.r1
    c0 = bounds.c0
    c1 = bounds.c1
  }

  const rows: unknown[][] = []
  for (let R = r0; R <= r1; R++) {
    const row: unknown[] = []
    for (let C = c0; C <= c1; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C })
      const cell = sheet[addr]
      row.push(cell?.v ?? "")
    }
    rows.push(row)
  }

  const maxRowsTotal = 1 + cap
  const truncated = rows.length > maxRowsTotal
  const sliced = truncated ? rows.slice(0, maxRowsTotal) : rows

  if (options.asJson) {
    if (sliced.length === 0) {
      return { asJson: true, jsonRows: [], truncated: false, rowLimit: cap }
    }
    const headers = sliced[0]!.map((h) => String(h ?? "").trim() || `col_${String(h)}`)
    const out: Record<string, unknown>[] = []
    for (let i = 1; i < sliced.length; i++) {
      const obj: Record<string, unknown> = {}
      headers.forEach((h, j) => {
        obj[h] = sliced[i]![j]
      })
      out.push(obj)
    }
    return { asJson: true, jsonRows: out, truncated, rowLimit: cap }
  }

  let md = sheetToMarkdown(sliced)
  if (truncated) {
    md += `\n\n_[Saída truncada: no máximo ${cap} linhas de dados (excl. cabeçalho). Ajuste maxRows ou previewOnly._`
  }
  return { asJson: false, markdown: md, truncated, rowLimit: cap }
}

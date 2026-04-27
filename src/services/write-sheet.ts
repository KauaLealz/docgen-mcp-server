import ExcelJS from "exceljs"
import fs from "node:fs"
import path from "node:path"
import { validateWritePath } from "../shared/security.js"

function csvEscape(cell: unknown): string {
  const s = String(cell ?? "")
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`
  return s
}

function writeCsv(outPath: string, headers: string[], rowsData: unknown[][]): void {
  const lines = [headers.map(csvEscape).join(","), ...rowsData.map((r) => r.map(csvEscape).join(","))]
  const bom = "\uFEFF"
  fs.writeFileSync(outPath, bom + lines.join("\r\n"), "utf8")
}

/** Extrai headers e linhas de valores a partir de data[]. */
function normalizeRows(params: {
  data: unknown[]
  columns?: Record<string, string>
}): { headers: string[]; rows: unknown[][] } {
  const rows = Array.isArray(params.data) ? params.data : []
  if (rows.length === 0) {
    return { headers: [], rows: [] }
  }
  const first = rows[0]
  let keys: string[]
  if (first && typeof first === "object" && !Array.isArray(first)) {
    keys = Object.keys(first as object)
  } else {
    keys =
      rows[0] && Array.isArray(rows[0])
        ? (rows[0] as unknown[]).map((_, i) => `col${i + 1}`)
        : ["value"]
  }

  const headers =
    params.columns && Object.keys(params.columns).length > 0
      ? keys.map((k) => params.columns![k] ?? k)
      : keys

  const rowValues: unknown[][] = []
  for (const row of rows) {
    if (row && typeof row === "object" && !Array.isArray(row)) {
      const obj = row as Record<string, unknown>
      rowValues.push(keys.map((k) => obj[k]))
    } else if (Array.isArray(row)) {
      rowValues.push(row)
    } else {
      rowValues.push([row])
    }
  }

  return { headers, rows: rowValues }
}

export async function writeSheet(params: {
  path: string
  data: unknown[]
  columns?: Record<string, string>
  freezePanes?: boolean
  append?: boolean
}): Promise<string> {
  const outPath = validateWritePath(params.path)
  const ext = path.extname(outPath).toLowerCase()

  const { headers, rows: rowValues } = normalizeRows(params)

  if (ext === ".csv") {
    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    if (rowValues.length === 0) {
      fs.writeFileSync(outPath, "\uFEFF", "utf8")
      return `CSV vazio gravado em ${outPath}.`
    }
    writeCsv(outPath, headers, rowValues)
    return `CSV gravado em ${outPath} (${rowValues.length} linha(s)).`
  }

  if (ext !== ".xlsx") {
    throw new Error("write_sheet: use extensão .xlsx ou .csv.")
  }

  const wb = new ExcelJS.Workbook()
  let ws: ExcelJS.Worksheet

  if (params.append) {
    if (ext !== ".xlsx") {
      throw new Error("append só é suportado para .xlsx.")
    }
  }

  if (params.append && fs.existsSync(outPath)) {
    await wb.xlsx.readFile(outPath)
    ws = wb.worksheets[0]!
    if (!ws) throw new Error("append: workbook sem folhas.")
    for (const rv of rowValues) {
      ws.addRow(rv)
    }
    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    await wb.xlsx.writeFile(outPath)
    return `Linhas anexadas em ${outPath} (${rowValues.length} linha(s)).`
  }

  ws = wb.addWorksheet("Sheet1")

  if (rowValues.length === 0) {
    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    await wb.xlsx.writeFile(outPath)
    return `Planilha vazia gravada em ${outPath}.`
  }

  ws.addRow(headers)

  for (const rv of rowValues) {
    ws.addRow(rv)
  }

  if (params.freezePanes) {
    ws.views = [{ state: "frozen", ySplit: 1 }]
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true })
  await wb.xlsx.writeFile(outPath)
  return `Planilha gravada em ${outPath} (${rowValues.length} linha(s) de dados + cabeçalho).`
}

import ExcelJS from "exceljs"
import fs from "node:fs"
import path from "node:path"
import { validatePath, validateWritePath } from "../shared/security.js"

export type CellUpdate = { cell: string; value: unknown; style?: Record<string, unknown> }

export async function patchSheet(
  filePath: string,
  updates: CellUpdate[],
): Promise<string> {
  const resolved = validatePath(filePath, true)
  const outPath = validateWritePath(resolved)
  const ext = path.extname(resolved).toLowerCase()
  if (ext !== ".xlsx") throw new Error("patch_sheet suporta apenas .xlsx nesta versão.")
  if (!updates.length) throw new Error("Informe ao menos uma atualização em updates.")

  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(resolved)
  const ws = wb.worksheets[0]
  if (!ws) throw new Error("Planilha sem folhas.")

  for (const u of updates) {
    const cell = ws.getCell(u.cell)
    cell.value = u.value as ExcelJS.CellValue
    if (u.style && typeof u.style === "object") {
      const s = u.style as {
        font?: Partial<ExcelJS.Font>
        fill?: ExcelJS.Fill
        border?: Partial<ExcelJS.Borders>
        alignment?: Partial<ExcelJS.Alignment>
        numFmt?: string
      }
      if (s.font) cell.font = { ...cell.font, ...s.font }
      if (s.fill) cell.fill = s.fill
      if (s.border) cell.border = { ...cell.border, ...s.border }
      if (s.alignment) cell.alignment = { ...cell.alignment, ...s.alignment }
      if (s.numFmt != null) cell.numFmt = s.numFmt
    }
  }

  fs.mkdirSync(path.dirname(outPath), { recursive: true })
  await wb.xlsx.writeFile(outPath)
  return `Planilha atualizada: ${outPath} (${updates.length} célula(s)).`
}

import fs from "node:fs"
import path from "node:path"
import { PDFDocument, rgb, StandardFonts } from "pdf-lib"
import JSZip from "jszip"
import { validatePath, validateWritePath } from "../shared/security.js"

async function loadPdf(buf: Buffer) {
  return PDFDocument.load(buf)
}

export async function patchDoc(params: {
  path: string
  action: "merge" | "split" | "watermark" | "replace_text"
  payload: unknown
}): Promise<string> {
  if (params.action === "merge") {
    const outPath = validateWritePath(params.path)
    const p = params.payload as { sources?: string[] }
    const sources = (p?.sources ?? []).map((s) => validatePath(s, true))
    if (sources.length === 0) throw new Error("merge: informe payload.sources (PDFs na ordem desejada).")
    const merged = await PDFDocument.create()
    for (const src of sources) {
      if (path.extname(src).toLowerCase() !== ".pdf") {
        throw new Error(`merge: apenas PDF suportado. Arquivo inválido: ${src}`)
      }
      const pdf = await loadPdf(fs.readFileSync(src))
      const pages = await merged.copyPages(pdf, pdf.getPageIndices())
      pages.forEach((page) => merged.addPage(page))
    }
    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    fs.writeFileSync(outPath, await merged.save())
    return `PDF mesclado gravado em ${outPath} (${sources.length} arquivo(s)).`
  }

  const resolved = validatePath(params.path, true)
  const ext = path.extname(resolved).toLowerCase()

  if (params.action === "split") {
    if (ext !== ".pdf") throw new Error("split: apenas arquivo .pdf na entrada.")
    const p = params.payload as { outputDir?: string }
    const outDir = p?.outputDir ? validateWritePath(p.outputDir) : path.join(path.dirname(resolved), `${path.basename(resolved, ".pdf")}_split`)
    fs.mkdirSync(outDir, { recursive: true })
    const pdf = await loadPdf(fs.readFileSync(resolved))
    const count = pdf.getPageCount()
    for (let i = 0; i < count; i++) {
      const one = await PDFDocument.create()
      const [copied] = await one.copyPages(pdf, [i])
      one.addPage(copied)
      const partPath = path.join(outDir, `page-${String(i + 1).padStart(3, "0")}.pdf`)
      fs.writeFileSync(partPath, await one.save())
    }
    return `PDF dividido em ${count} arquivo(s) em ${outDir}.`
  }

  if (params.action === "watermark") {
    if (ext !== ".pdf") throw new Error("watermark: apenas .pdf suportado nesta versão.")
    const p = params.payload as { text?: string; opacity?: number }
    const text = p?.text ?? "CONFIDENCIAL"
    const opacity = p?.opacity ?? 0.15
    const pdf = await loadPdf(fs.readFileSync(resolved))
    const font = await pdf.embedFont(StandardFonts.HelveticaBold)
    const pages = pdf.getPages()
    for (const page of pages) {
      const { width, height } = page.getSize()
      page.drawText(text, {
        x: width / 4,
        y: height / 2,
        size: Math.min(width, height) / 18,
        font,
        color: rgb(0.6, 0.6, 0.6),
        opacity,
      })
    }
    const outPath = validateWritePath(resolved)
    fs.writeFileSync(outPath, await pdf.save())
    return `Marca d'água aplicada em ${outPath}.`
  }

  if (params.action === "replace_text") {
    const p = params.payload as { replacements?: { from: string; to: string }[] }
    const reps = p?.replacements ?? []
    if (reps.length === 0) throw new Error("replace_text: informe payload.replacements.")

    if (ext === ".docx") {
      const zip = await JSZip.loadAsync(fs.readFileSync(resolved))
      const xmlPath = "word/document.xml"
      const file = zip.file(xmlPath)
      if (!file) throw new Error("DOCX inválido: falta word/document.xml")
      let xml = await file.async("string")
      for (const r of reps) {
        xml = xml.split(r.from).join(r.to)
      }
      zip.file(xmlPath, xml)
      const outBuf = await zip.generateAsync({ type: "nodebuffer" })
      const outPath = validateWritePath(resolved)
      fs.writeFileSync(outPath, Buffer.from(outBuf))
      return `Substituições aplicadas no DOCX em ${outPath}.`
    }

    if (ext === ".pdf") {
      throw new Error(
        "replace_text em PDF não é suportado de forma confiável nesta versão (texto pode ser vetorizado). Exporte para DOCX ou use ferramentas dedicadas.",
      )
    }

    throw new Error(`replace_text: extensão não suportada: ${ext}`)
  }

  throw new Error(`Ação desconhecida: ${params.action}`)
}

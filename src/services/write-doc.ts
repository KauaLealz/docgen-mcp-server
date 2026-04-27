import fs from "node:fs"
import path from "node:path"
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx"
import { PDFDocument, StandardFonts, rgb } from "pdf-lib"
import JSZip from "jszip"
import { markdownToDocxParagraphs, markdownToPdfPlainLines } from "../shared/markdown-docx.js"
import { validateWritePath } from "../shared/security.js"

function splitPlainParagraphs(content: string): Paragraph[] {
  const blocks = content.split(/\n\s*\n/).filter((b) => b.trim())
  const out: Paragraph[] = []
  for (const block of blocks) {
    const lines = block.split("\n")
    const head = lines[0]?.trim() ?? ""
    const rest = lines.slice(1).join("\n").trim()
    if (head.startsWith("# ")) {
      out.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun(head.slice(2))],
        }),
      )
      continue
    }
    if (head.startsWith("## ")) {
      out.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun(head.slice(3))],
        }),
      )
      continue
    }
    const text = rest ? `${head}\n${rest}` : head
    for (const line of text.split("\n")) {
      out.push(new Paragraph({ children: [new TextRun(line)] }))
    }
  }
  return out
}

async function mergeFieldsIntoDocxTemplate(
  templateBuf: Buffer,
  mergeFields: Record<string, unknown>,
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(templateBuf)
  const xmlPath = "word/document.xml"
  const file = zip.file(xmlPath)
  if (!file) throw new Error("Template DOCX inválido: falta word/document.xml")
  let xml = await file.async("string")
  for (const [k, v] of Object.entries(mergeFields)) {
    const token = `{{${k}}}`
    xml = xml.split(token).join(String(v ?? ""))
  }
  zip.file(xmlPath, xml)
  const out = await zip.generateAsync({ type: "nodebuffer" })
  return Buffer.from(out)
}

export async function writeDoc(params: {
  path: string
  type: "docx" | "pdf"
  content: string
  templatePath?: string
  mergeFields?: Record<string, unknown>
  contentFormat?: "markdown" | "plain"
}): Promise<string> {
  const outPath = validateWritePath(params.path)
  const fmt = params.contentFormat ?? "markdown"

  if (params.type === "docx") {
    let buffer: Buffer
    if (params.templatePath) {
      const tpl = validateWritePath(params.templatePath)
      if (!fs.existsSync(tpl)) throw new Error(`Template não encontrado: ${params.templatePath}`)
      const tplBuf = fs.readFileSync(tpl)
      buffer =
        params.mergeFields && Object.keys(params.mergeFields).length > 0
          ? await mergeFieldsIntoDocxTemplate(tplBuf, params.mergeFields)
          : tplBuf
      fs.mkdirSync(path.dirname(outPath), { recursive: true })
      fs.writeFileSync(outPath, buffer)
      return `DOCX gravado em ${outPath}${params.mergeFields ? " (template + mergeFields)." : " (cópia do template)."}.`
    }

    const paras =
      fmt === "markdown"
        ? markdownToDocxParagraphs(params.content)
        : splitPlainParagraphs(params.content)
    const doc = new Document({
      sections: [
        {
          children:
            paras.length > 0
              ? paras
              : [new Paragraph({ children: [new TextRun(params.content)] })],
        },
      ],
    })
    buffer = Buffer.from(await Packer.toBuffer(doc))
    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    fs.writeFileSync(outPath, buffer)
    return `DOCX gravado em ${outPath}${fmt === "markdown" ? " (Markdown)." : "."}`
  }

  if (params.type === "pdf") {
    const pdf = await PDFDocument.create()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    const pageSize: [number, number] = [595.28, 841.89]
    let page = pdf.addPage(pageSize)
    const fontSize = 11
    const margin = 50
    let y = pageSize[1] - margin
    const maxW = pageSize[0] - margin * 2
    const lines =
      fmt === "markdown" ? markdownToPdfPlainLines(params.content) : params.content.split(/\r?\n/)

    const drawLine = (text: string) => {
      const wrapped = wrapText(text, font, fontSize, maxW)
      for (const line of wrapped) {
        if (y < margin + fontSize) {
          page = pdf.addPage(pageSize)
          y = pageSize[1] - margin
        }
        page.drawText(line, { x: margin, y, size: fontSize, font, color: rgb(0, 0, 0) })
        y -= fontSize + 2
      }
    }

    for (const line of lines) {
      drawLine(line.length ? line : " ")
    }

    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    fs.writeFileSync(outPath, await pdf.save())
    return `PDF gravado em ${outPath}${fmt === "markdown" ? " (Markdown)." : "."}`
  }

  throw new Error(`Tipo não suportado: ${params.type}`)
}

function wrapText(
  text: string,
  font: { widthOfTextAtSize: (t: string, s: number) => number },
  size: number,
  maxW: number,
): string[] {
  const words = text.split(/\s+/)
  const lines: string[] = []
  let cur = ""
  for (const w of words) {
    const trial = cur ? `${cur} ${w}` : w
    if (font.widthOfTextAtSize(trial, size) <= maxW) {
      cur = trial
    } else {
      if (cur) lines.push(cur)
      cur = w
    }
  }
  if (cur) lines.push(cur)
  return lines.length ? lines : [""]
}

import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { randomBytes } from "node:crypto"
import JSZip from "jszip"
import mammoth from "mammoth"
import { XMLParser } from "fast-xml-parser"
import TurndownService from "turndown"
import { validatePath } from "../shared/security.js"

const turndown = new TurndownService({ headingStyle: "atx" })

async function readPdfBuffer(buf: Buffer): Promise<string> {
  const pdfParse = (await import("pdf-parse")).default as (b: Buffer) => Promise<{ text: string }>
  const data = await pdfParse(buf)
  return (data.text ?? "").trim()
}

function odtXmlToPlain(xml: string): string {
  const parser = new XMLParser({
    ignoreAttributes: false,
    removeNSPrefix: true,
    trimValues: true,
  })
  const obj = parser.parse(xml)
  const collect = (v: unknown): string[] => {
    if (v == null) return []
    if (typeof v === "string") return v.trim() ? [v] : []
    if (Array.isArray(v)) return v.flatMap(collect)
    if (typeof v === "object") return Object.values(v).flatMap(collect)
    return []
  }
  return collect(obj).join("\n\n").trim()
}

async function readOdt(pathResolved: string): Promise<string> {
  const buf = fs.readFileSync(pathResolved)
  const zip = await JSZip.loadAsync(buf)
  const entry = zip.file("content.xml")
  if (!entry) throw new Error("content.xml não encontrado no ODT.")
  const xml = await entry.async("string")
  return odtXmlToPlain(xml)
}

export type ReadDocResult = {
  markdown: string
  imagePaths: string[]
  note?: string
}

function truncateMarkdown(
  text: string,
  maxChars?: number,
  previewOnly?: boolean,
): { text: string; truncated: boolean } {
  const previewCap = 12_000
  const cap = maxChars ?? (previewOnly ? previewCap : undefined)
  if (cap == null || text.length <= cap) return { text, truncated: false }
  return {
    text: `${text.slice(0, cap)}\n\n_[… texto truncado; use maxChars ou desative previewOnly …]_`,
    truncated: true,
  }
}

export async function readDoc(
  filePath: string,
  options: {
    includeComments?: boolean
    extractImages?: boolean
    maxChars?: number
    previewOnly?: boolean
  },
): Promise<ReadDocResult & { truncated?: boolean }> {
  const resolved = validatePath(filePath, true)
  const ext = path.extname(resolved).toLowerCase()
  const buf = fs.readFileSync(resolved)

  let note: string | undefined
  if (options.includeComments) {
    note =
      "includeComments: comentários em Word/PDF podem não ser extraídos integralmente nesta versão."
  }

  if (ext === ".pdf") {
    const text = await readPdfBuffer(buf)
    let md = text || "(PDF sem texto extraível.)"
    const tr = truncateMarkdown(md, options.maxChars, options.previewOnly)
    md = tr.text
    return { markdown: md, imagePaths: [], note, truncated: tr.truncated }
  }

  if (ext === ".odt") {
    const text = await readOdt(resolved)
    let md = text || "(ODT vazio.)"
    const tr = truncateMarkdown(md, options.maxChars, options.previewOnly)
    md = tr.text
    return { markdown: md, imagePaths: [], note, truncated: tr.truncated }
  }

  if (ext === ".docx") {
    const imagePaths: string[] = []

    if (!options.extractImages) {
      const raw = await mammoth.extractRawText({ buffer: buf })
      let md = (raw.value ?? "").trim()
      md = md || "(DOCX sem texto.)"
      const tr = truncateMarkdown(md, options.maxChars, options.previewOnly)
      return {
        markdown: tr.text,
        imagePaths,
        note,
        truncated: tr.truncated,
      }
    }

    const imageDir = path.join(os.tmpdir(), `docgen-img-${randomBytes(8).toString("hex")}`)
    fs.mkdirSync(imageDir, { recursive: true })

    const result = await mammoth.convertToHtml(
      { buffer: buf },
      {
        convertImage: mammoth.images.imgElement(async (image) => {
          const ab = await image.read("arraybuffer")
          const b = Buffer.from(ab)
          const extImg = image.contentType?.includes("png")
            ? ".png"
            : image.contentType?.includes("jpeg")
              ? ".jpg"
              : ".bin"
          const name = `image-${imagePaths.length + 1}${extImg}`
          const out = path.join(imageDir, name)
          fs.writeFileSync(out, b)
          imagePaths.push(out)
          return { src: out }
        }),
      },
    )

    let md = turndown.turndown(result.value || "").trim()
    md = md || "(DOCX sem texto.)"
    const tr = truncateMarkdown(md, options.maxChars, options.previewOnly)
    return {
      markdown: tr.text,
      imagePaths,
      note,
      truncated: tr.truncated,
    }
  }

  throw new Error(`Extensão não suportada para read_doc: ${ext}. Use .docx, .pdf ou .odt.`)
}

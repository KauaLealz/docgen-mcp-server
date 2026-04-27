import { HeadingLevel, Paragraph, TextRun } from "docx"

function headingLevelFromHashes(n: number): (typeof HeadingLevel)[keyof typeof HeadingLevel] {
  if (n <= 1) return HeadingLevel.HEADING_1
  if (n === 2) return HeadingLevel.HEADING_2
  if (n === 3) return HeadingLevel.HEADING_3
  if (n === 4) return HeadingLevel.HEADING_4
  if (n === 5) return HeadingLevel.HEADING_5
  return HeadingLevel.HEADING_6
}

/**
 * Markdown comum → parágrafos docx: títulos (#–######), listas (-/*), blocos ```, parágrafos.
 */
export function markdownToDocxParagraphs(md: string): Paragraph[] {
  const lines = md.replace(/\r\n/g, "\n").split("\n")
  const out: Paragraph[] = []
  let i = 0
  while (i < lines.length) {
    const raw = lines[i]!
    const t = raw.trim()
    if (!t) {
      i++
      continue
    }

    if (t.startsWith("```")) {
      i++
      const code: string[] = []
      while (i < lines.length && !lines[i]!.trim().startsWith("```")) {
        code.push(lines[i]!)
        i++
      }
      if (i < lines.length) i++
      for (const line of code) {
        out.push(
          new Paragraph({
            children: [new TextRun({ font: "Consolas", text: line })],
            spacing: { before: 40, after: 40 },
          }),
        )
      }
      continue
    }

    const hm = t.match(/^(#{1,6})\s+(.*)$/)
    if (hm) {
      const level = hm[1]!.length
      const text = hm[2]!.trim()
      out.push(
        new Paragraph({
          heading: headingLevelFromHashes(level),
          children: [new TextRun(text)],
        }),
      )
      i++
      continue
    }

    if (/^[\-\*]\s+/.test(t)) {
      const items: string[] = []
      while (i < lines.length) {
        const lt = lines[i]!.trim()
        const im = lt.match(/^[\-\*]\s+(.*)$/)
        if (!im) break
        items.push(im[1]!.trim())
        i++
      }
      for (const item of items) {
        out.push(new Paragraph({ children: [new TextRun(`• ${item}`)] }))
      }
      continue
    }

    const paraLines: string[] = [raw]
    i++
    while (i < lines.length && lines[i]!.trim()) {
      const nt = lines[i]!.trim()
      if (/^(#{1,6}\s|```|[\-\*]\s)/.test(nt)) break
      paraLines.push(lines[i]!)
      i++
    }
    const body = paraLines.join("\n").trim()
    if (body) {
      for (const pl of body.split("\n")) {
        out.push(new Paragraph({ children: [new TextRun(pl)] }))
      }
    }
  }
  return out
}

/** PDF simples a partir de Markdown: prefixos visuais por nível de título. */
export function markdownToPdfPlainLines(md: string): string[] {
  const lines = md.replace(/\r\n/g, "\n").split("\n")
  const out: string[] = []
  let i = 0
  while (i < lines.length) {
    const t = lines[i]!.trim()
    if (!t) {
      out.push("")
      i++
      continue
    }
    if (t.startsWith("```")) {
      i++
      while (i < lines.length && !lines[i]!.trim().startsWith("```")) {
        out.push(`  ${lines[i]}`)
        i++
      }
      if (i < lines.length) i++
      continue
    }
    const hm = t.match(/^(#{1,6})\s+(.*)$/)
    if (hm) {
      const level = hm[1]!.length
      const text = hm[2]!.trim()
      out.push(`${"#".repeat(Math.min(level, 3))} ${text}`)
      i++
      continue
    }
    if (/^[\-\*]\s+/.test(t)) {
      while (i < lines.length) {
        const lt = lines[i]!.trim()
        const im = lt.match(/^[\-\*]\s+(.*)$/)
        if (!im) break
        out.push(`• ${im[1]}`)
        i++
      }
      continue
    }
    out.push(lines[i]!)
    i++
  }
  return out
}

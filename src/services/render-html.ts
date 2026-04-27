import fs from "node:fs"
import path from "node:path"
import archiver from "archiver"
import puppeteer from "puppeteer"
import { validateWritePath } from "../shared/security.js"

function wrapHtml(body: string, css: string): string {
  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"/><style>${css}</style></head><body>${body}</body></html>`
}

export async function renderSlide(params: {
  html: string
  css: string
  format: "pdf" | "zip"
  aspectRatio: "16:9" | "4:3"
  outputPath: string
}): Promise<string> {
  const outPath = validateWritePath(params.outputPath)
  const w = params.aspectRatio === "16:9" ? 1920 : 1024
  const h = params.aspectRatio === "16:9" ? 1080 : 768

  const slideCss = `
    @page { size: ${w}px ${h}px; margin: 0; }
    html, body { margin: 0; padding: 0; background: #fff; }
    .slide {
      width: ${w}px;
      min-height: ${h}px;
      box-sizing: border-box;
      page-break-after: always;
      page-break-inside: avoid;
    }
    .slide:last-child { page-break-after: auto; }
    ${params.css}
  `
  const full = wrapHtml(params.html, slideCss)

  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  })
  try {
    const page = await browser.newPage()
    await page.setViewport({ width: w, height: h, deviceScaleFactor: 1 })
    await page.setContent(full, { waitUntil: "networkidle0" })

    if (params.format === "pdf") {
      const dir = path.dirname(outPath)
      fs.mkdirSync(dir, { recursive: true })
      await page.pdf({
        path: outPath,
        width: `${w}px`,
        height: `${h}px`,
        printBackground: true,
        preferCSSPageSize: true,
      })
      return `Slides exportados para PDF: ${outPath}`
    }

    const slides = await page.$$eval(".slide", (els) =>
      els.map((e) => (e as HTMLElement).outerHTML),
    )
    const htmlParts = slides.length > 0 ? slides : [params.html]

    const zipPath = outPath.toLowerCase().endsWith(".zip") ? outPath : `${outPath}.zip`
    fs.mkdirSync(path.dirname(zipPath), { recursive: true })

    const archive = archiver("zip", { zlib: { level: 9 } })
    const out = fs.createWriteStream(zipPath)
    await new Promise<void>((resolve, reject) => {
      archive.on("error", reject)
      out.on("close", () => resolve())
      archive.pipe(out)
      htmlParts.forEach((chunk, i) => {
        const doc = wrapHtml(chunk, `${slideCss}\n${params.css}`)
        archive.append(doc, { name: `slide-${String(i + 1).padStart(3, "0")}.html` })
      })
      archive.append(Buffer.from(slideCss + "\n" + params.css, "utf8"), {
        name: "styles-reference.css",
      })
      archive.finalize()
    })
    return `Slides exportados como ZIP: ${zipPath}`
  } finally {
    await browser.close()
  }
}

export async function renderPage(params: {
  html: string
  css: string
  outputPath: string
  generateTOC?: boolean
  margins?: { top?: string; right?: string; bottom?: string; left?: string }
}): Promise<string> {
  const outPath = validateWritePath(params.outputPath)
  let tocCss = ""
  let bodyExtra = ""
  if (params.generateTOC) {
    tocCss = `
      #docgen-toc { page-break-after: always; font-family: system-ui, sans-serif; }
      #docgen-toc h2 { font-size: 14pt; }
      #docgen-toc nav { margin-top: 12px; }
      #docgen-toc a { color: inherit; text-decoration: none; }
    `
    bodyExtra = `
      <div id="docgen-toc"><h2>Índice</h2><nav id="toc-nav"></nav></div>
      <script>
        document.addEventListener('DOMContentLoaded', () => {
          const nav = document.getElementById('toc-nav');
          const hs = document.querySelectorAll('h1, h2, h3');
          hs.forEach((h, i) => {
            if (!h.id) h.id = 'heading-' + i;
            const a = document.createElement('a');
            a.href = '#' + h.id;
            a.textContent = h.textContent || '';
            const p = document.createElement('p');
            p.appendChild(a);
            nav?.appendChild(p);
          });
        });
      </script>
    `
  }

  const margin = params.margins ?? {}
  const top = margin.top ?? "20mm"
  const right = margin.right ?? "15mm"
  const bottom = margin.bottom ?? "20mm"
  const left = margin.left ?? "15mm"
  const full = wrapHtml(params.html + bodyExtra, `${tocCss}\n${params.css}`)

  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  })
  try {
    const page = await browser.newPage()
    await page.setContent(full, { waitUntil: "networkidle0" })
    fs.mkdirSync(path.dirname(outPath), { recursive: true })
    await page.pdf({
      path: outPath,
      format: "A4",
      printBackground: true,
      margin: { top, right, bottom, left },
      displayHeaderFooter: false,
    })
    return `Documento PDF gerado: ${outPath}`
  } finally {
    await browser.close()
  }
}

import fs from "node:fs"
import path from "node:path"
import fg from "fast-glob"
import { getScanMaxMatches } from "../config/limits.js"
import yauzl from "../shared/yauzl-cjs.js"
import { validatePath } from "../shared/security.js"

async function scanZip(zipPath: string, query: RegExp): Promise<string[]> {
  const hits: string[] = []
  await new Promise<void>((resolve, reject) => {
    yauzl.open(zipPath, { lazyEntries: true }, (err, zipfile) => {
      if (err || !zipfile) {
        reject(err ?? new Error("ZIP inválido"))
        return
      }
      zipfile.readEntry()
      zipfile.on("entry", (entry) => {
        const name = entry.fileName.replace(/\\/g, "/")
        if (!entry.isDirectory && query.test(name)) {
          hits.push(`${zipPath}::${name}`)
        }
        zipfile.readEntry()
      })
      zipfile.on("end", () => resolve())
      zipfile.on("error", reject)
    })
  })
  return hits
}

async function readFileSnippet(filePath: string, max = 8000): Promise<string> {
  const fd = fs.openSync(filePath, "r")
  try {
    const buf = Buffer.alloc(Math.min(max, fs.statSync(filePath).size))
    fs.readSync(fd, buf, 0, buf.length, 0)
    return buf.toString("utf8")
  } finally {
    fs.closeSync(fd)
  }
}

export async function scanDir(
  rootPath: string,
  query: string,
  recursive: boolean,
  maxMatches?: number,
): Promise<{ matches: string[]; truncated: boolean }> {
  const resolved = validatePath(rootPath, true)
  let re: RegExp
  try {
    re = new RegExp(query, "ims")
  } catch {
    throw new Error("query não é uma expressão regular válida.")
  }

  const st = fs.statSync(resolved)
  const cap = maxMatches != null && maxMatches > 0 ? maxMatches : getScanMaxMatches()

  const trim = (arr: string[]): { matches: string[]; truncated: boolean } => {
    if (arr.length <= cap) return { matches: arr, truncated: false }
    return { matches: arr.slice(0, cap), truncated: true }
  }

  if (st.isFile() && resolved.toLowerCase().endsWith(".zip")) {
    const m = await scanZip(resolved, re)
    return trim(m)
  }

  if (st.isFile()) {
    const text = await readFileSnippet(resolved)
    const lines = text.split(/\n/)
    const hits: string[] = []
    lines.forEach((line, i) => {
      if (re.test(line)) hits.push(`${resolved}:${i + 1}:${line.slice(0, 200)}`)
    })
    return trim(hits)
  }

  const pattern = recursive ? "**/*" : "*"
  const files = await fg(pattern, {
    cwd: resolved,
    onlyFiles: true,
    dot: false,
    absolute: true,
  })

  const matches: string[] = []
  for (const file of files) {
    const base = path.basename(file)
    if (re.test(base)) {
      matches.push(file)
      continue
    }
    if (file.toLowerCase().endsWith(".zip")) {
      matches.push(...(await scanZip(file, re)))
      continue
    }
    try {
      const text = await readFileSnippet(file, 256 * 1024)
      if (re.test(text)) matches.push(file)
    } catch {
      /* binário ou inacessível */
    }
  }

  return trim(matches)
}

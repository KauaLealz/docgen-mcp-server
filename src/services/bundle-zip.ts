import fs from "node:fs"
import path from "node:path"
import archiver from "archiver"
import { validatePath, validateWritePath } from "../shared/security.js"

export async function bundleZip(files: string[], outputName: string): Promise<string> {
  if (!files.length) throw new Error("Informe ao menos um arquivo em files.")
  const resolved = files.map((f) => validatePath(f, true))
  const outPathRaw = outputName.includes("/") || outputName.includes("\\") ? outputName : path.resolve(process.cwd(), outputName)
  const outPath = validateWritePath(outPathRaw)
  fs.mkdirSync(path.dirname(outPath), { recursive: true })

  await new Promise<void>((resolve, reject) => {
    const output = fs.createWriteStream(outPath)
    const archive = archiver("zip", { zlib: { level: 9 } })
    output.on("close", () => resolve())
    archive.on("error", reject)
    archive.pipe(output)
    for (const f of resolved) {
      const name = path.basename(f)
      archive.file(f, { name })
    }
    archive.finalize()
  })

  return `Arquivo ZIP criado: ${outPath} (${files.length} arquivo(s)).`
}

import { existsSync, readFileSync, writeFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

const root = join(dirname(fileURLToPath(import.meta.url)), "..")
const arg = process.argv[2]
const level =
  arg === "minor" || arg === "major" ? arg : "patch"

function nextVersion(current) {
  const m = current.match(/^(\d+)\.(\d+)\.(\d+)$/)
  if (!m) {
    throw new Error(`Versão inválida no package.json (esperado X.Y.Z): ${current}`)
  }
  let x = Number(m[1])
  let y = Number(m[2])
  let z = Number(m[3])
  if (level === "major") {
    x++
    y = 0
    z = 0
  } else if (level === "minor") {
    y++
    z = 0
  } else {
    z++
  }
  return `${x}.${y}.${z}`
}

const pkgPath = join(root, "package.json")
const pkg = JSON.parse(readFileSync(pkgPath, "utf8"))
const nv = nextVersion(pkg.version)
pkg.version = nv
writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + "\n")

const lockPath = join(root, "package-lock.json")
if (existsSync(lockPath)) {
  const lock = JSON.parse(readFileSync(lockPath, "utf8"))
  lock.version = nv
  if (lock.packages?.[""]) {
    lock.packages[""].version = nv
  }
  writeFileSync(lockPath, JSON.stringify(lock, null, 2) + "\n")
}

process.stdout.write(`${nv}\n`)

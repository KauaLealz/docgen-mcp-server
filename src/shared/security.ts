import fs from "node:fs"
import os from "node:os"
import path from "node:path"

export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024

const BLOCKED_DIR_NAMES = new Set([
  ".ssh",
  ".aws",
  ".gnupg",
  ".credentials",
  ".kube",
  ".docker",
  ".azure",
])

function getBlockedRoots(): string[] {
  const blocked: string[] = []
  const isWin = process.platform === "win32"
  if (isWin) {
    const sysRoot = process.env.SystemRoot ?? "C:\\Windows"
    blocked.push(path.resolve(sysRoot))
    blocked.push(path.resolve("C:/Program Files"))
    blocked.push(path.resolve("C:/Program Files (x86)"))
  } else {
    for (const p of ["/bin", "/sbin", "/usr", "/etc", "/boot", "/sys", "/proc"]) {
      try {
        blocked.push(path.resolve(p))
      } catch {
        /* ignore */
      }
    }
  }
  const home = os.homedir()
  for (const name of BLOCKED_DIR_NAMES) {
    blocked.push(path.resolve(path.join(home, name)))
  }
  return blocked
}

function getAllowedRoots(): string[] {
  const raw = process.env.DOCGEN_ALLOWED_ROOTS?.trim()
  if (!raw) return []
  return raw
    .split(",")
    .map((s) => path.normalize(path.resolve(s.trim())))
    .filter(Boolean)
}

function enforceAllowedRoots(resolved: string): void {
  const roots = getAllowedRoots()
  if (roots.length === 0) return
  const norm = path.normalize(resolved)
  const ok = roots.some((r) => norm === r || norm.startsWith(r + path.sep))
  if (!ok) {
    throw new Error(
      `Caminho fora de DOCGEN_ALLOWED_ROOTS (defina caminhos absolutos separados por vírgula). Recebido: ${resolved}`,
    )
  }
}

export function validatePath(filePath: string, mustExist: boolean): string {
  if (filePath.includes("..")) {
    throw new Error(`Path traversal detectado: ${filePath}`)
  }
  const resolved = path.resolve(filePath)
  enforceAllowedRoots(resolved)
  if (mustExist && !fs.existsSync(resolved)) {
    throw new Error(`Arquivo não encontrado: ${filePath}`)
  }
  if (mustExist) {
    const st = fs.statSync(resolved)
    if (st.isFile() && st.size > MAX_FILE_SIZE_BYTES) {
      throw new Error(
        `Arquivo excede limite de ${MAX_FILE_SIZE_BYTES / (1024 * 1024)}MB: ${(st.size / (1024 * 1024)).toFixed(1)}MB`,
      )
    }
  }
  return resolved
}

export function validateWritePath(filePath: string): string {
  const resolved = validatePath(filePath, false)
  const normalized = path.normalize(resolved)
  for (const b of getBlockedRoots()) {
    const nb = path.normalize(b)
    if (normalized === nb || normalized.startsWith(nb + path.sep)) {
      throw new Error(`Escrita bloqueada em diretório protegido: ${b}`)
    }
  }
  return resolved
}

/** Valida vários caminhos de leitura; lança se algum inválido. */
export function validateReadPaths(paths: string[], mustExist = true): string[] {
  return paths.map((p) => validatePath(p, mustExist))
}

/** Limites configuráveis por ambiente (sem over-engineering). */

export function getScanMaxMatches(): number {
  const raw = process.env.DOCGEN_SCAN_MAX_MATCHES
  const n = raw ? parseInt(raw, 10) : 500
  if (!Number.isFinite(n) || n < 1) return 500
  return Math.min(n, 50_000)
}

/** Máximo de linhas de dados lidas da folha (após cabeçalho JSON ou linhas Markdown). */
export function getReadSheetMaxRows(): number {
  const raw = process.env.DOCGEN_READ_SHEET_MAX_ROWS
  const n = raw ? parseInt(raw, 10) : 10_000
  if (!Number.isFinite(n) || n < 1) return 10_000
  return Math.min(n, 100_000)
}

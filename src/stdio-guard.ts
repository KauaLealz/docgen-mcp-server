import util from "node:util"

for (const m of ["log", "info", "debug", "warn"] as const) {
  console[m] = (...args: unknown[]) => {
    process.stderr.write(util.format(...args) + "\n")
  }
}

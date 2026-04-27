import { spawnSync } from "node:child_process"
import { fileURLToPath } from "node:url"
import { dirname, join } from "node:path"

const root = join(dirname(fileURLToPath(import.meta.url)), "..")

const tsc = spawnSync("tsc", ["--project", join(root, "tsconfig.json")], {
  cwd: root,
  stdio: "inherit",
  shell: true,
})

process.exit(tsc.status ?? 1)

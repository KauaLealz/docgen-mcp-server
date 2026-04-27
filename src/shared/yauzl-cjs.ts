import { createRequire } from "node:module"

/** yauzl é CommonJS; carregar via require evita falhas de resolução ESM em alguns ambientes (ex.: npx no Windows). */
const require = createRequire(import.meta.url)
export default require("yauzl") as typeof import("yauzl")

import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { registerReadTools } from "./register-read-tools.js"
import { registerWriteTools } from "./register-write-tools.js"
import { registerRenderTools } from "./register-render-tools.js"
import { registerPatchTools } from "./register-patch-tools.js"
import { registerSystemTools } from "./register-system-tools.js"

export function registerAllTools(server: McpServer): void {
  registerReadTools(server)
  registerWriteTools(server)
  registerRenderTools(server)
  registerPatchTools(server)
  registerSystemTools(server)
}

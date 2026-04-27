import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js"
import { ErrorService } from "./error-service.js"
import {
  errorResult,
  toCallToolResult,
  type ToolSuccessPayload,
} from "./tool-result.js"

export async function safeTool(
  tool: string,
  fn: () => Promise<ToolSuccessPayload>,
): Promise<CallToolResult> {
  try {
    return toCallToolResult(await fn())
  } catch (e) {
    const info = ErrorService.handle(e, tool)
    return errorResult({
      ok: false,
      code: info.type,
      tool,
      message: info.user_message,
      hint: ErrorService.hintFor(info, e),
    })
  }
}

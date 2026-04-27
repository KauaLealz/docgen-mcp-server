import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js"

export function textResult(text: string): CallToolResult {
  return { content: [{ type: "text", text }] }
}

export function textAndStructured(
  text: string,
  structured: Record<string, unknown>,
): CallToolResult {
  return {
    content: [{ type: "text", text }],
    structuredContent: structured,
  }
}

export type ToolSuccessPayload =
  | string
  | { text: string; structured: Record<string, unknown> }

export function toCallToolResult(payload: ToolSuccessPayload): CallToolResult {
  if (typeof payload === "string") return textResult(payload)
  return textAndStructured(payload.text, payload.structured)
}

export type StructuredErrorPayload = {
  ok: false
  code: string
  tool: string
  message: string
  hint?: string
}

export function errorResult(info: StructuredErrorPayload): CallToolResult {
  const text =
    info.hint != null && info.hint.length > 0
      ? `${info.message}\n\nDica: ${info.hint}`
      : info.message
  return {
    content: [{ type: "text", text }],
    structuredContent: info as unknown as Record<string, unknown>,
    isError: true,
  }
}

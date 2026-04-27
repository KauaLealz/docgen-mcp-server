export type ErrorInfo = {
  type: string
  message: string
  context: string
  user_message: string
}

export class ErrorService {
  static handle(error: unknown, context = ""): ErrorInfo {
    if (error instanceof Error) {
      const message = error.message
      if (error.name === "TimeoutError") {
        return {
          type: "TIMEOUT_ERROR",
          message,
          context,
          user_message:
            "A operação excedeu o tempo limite. Tente novamente ou reduza o tamanho do arquivo.",
        }
      }
      if (error instanceof TypeError || error instanceof SyntaxError) {
        return {
          type: "VALUE_ERROR",
          message,
          context,
          user_message: `Parâmetro inválido: ${message}`,
        }
      }
      if (
        message.includes("ENOENT") ||
        message.includes("no such file") ||
        message.toLowerCase().includes("cannot find")
      ) {
        return {
          type: "FILE_NOT_FOUND",
          message,
          context,
          user_message: `Arquivo ou pasta não encontrado: ${message}`,
        }
      }
      if (message.includes("EACCES") || message.includes("EPERM")) {
        return {
          type: "PERMISSION_ERROR",
          message,
          context,
          user_message: "Sem permissão para ler ou gravar neste caminho.",
        }
      }
      return {
        type: "ERROR",
        message,
        context,
        user_message: message || "Erro ao processar a solicitação.",
      }
    }
    return {
      type: "ERROR",
      message: String(error),
      context,
      user_message: String(error),
    }
  }

  /** Dica curta para structuredContent (DX). */
  static hintFor(info: ErrorInfo, error: unknown): string | undefined {
    if (info.type === "VALUE_ERROR" && error instanceof Error) {
      if (error.message.includes("range")) {
        return 'Use range no formato Excel, ex.: A1:D10 (colunas A–ZZ, linhas ≥ 1).'
      }
    }
    if (info.type === "ERROR" && error instanceof Error) {
      if (error.message.includes("DOCGEN_ALLOWED_ROOTS")) {
        return 'Defina DOCGEN_ALLOWED_ROOTS com caminhos absolutos separados por vírgula, ou remova a variável.'
      }
    }
    return undefined
  }

  static formatForAgent(info: ErrorInfo, includeDetails = false): string {
    let msg = info.user_message || info.message || "Erro desconhecido."
    if (includeDetails && info.context) {
      msg += ` (Contexto: ${info.context})`
    }
    return msg
  }
}

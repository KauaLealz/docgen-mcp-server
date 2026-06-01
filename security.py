import os
import sys

MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024

BLOCKED_DIR_NAMES = {
    ".ssh", ".aws", ".gnupg", ".credentials", ".kube", ".docker", ".azure"
}

def get_blocked_roots() -> list:
    blocked = []
    is_win = sys.platform == "win32"
    if is_win:
        sys_root = os.environ.get("SystemRoot", "C:\\Windows")
        blocked.append(os.path.abspath(sys_root))
        blocked.append(os.path.abspath("C:/Program Files"))
        blocked.append(os.path.abspath("C:/Program Files (x86)"))
    else:
        for p in ["/bin", "/sbin", "/usr", "/etc", "/boot", "/sys", "/proc"]:
            try:
                blocked.append(os.path.abspath(p))
            except Exception:
                pass
    home = os.path.expanduser("~")
    for name in BLOCKED_DIR_NAMES:
        blocked.append(os.path.abspath(os.path.join(home, name)))
    return blocked

def get_allowed_roots() -> list:
    raw = os.environ.get("DOCGEN_ALLOWED_ROOTS", "").strip()
    if not raw:
        return []
    return [os.path.abspath(s.strip()) for s in raw.split(",") if s.strip()]

def enforce_allowed_roots(resolved: str) -> None:
    roots = get_allowed_roots()
    if not roots:
        return
    norm = os.path.normpath(resolved)
    ok = False
    for r in roots:
        if norm == r or norm.startswith(r + os.sep):
            ok = True
            break
    if not ok:
        raise ValueError(f"Caminho fora de DOCGEN_ALLOWED_ROOTS. Recebido: {resolved}")

def validate_path(file_path: str, must_exist: bool) -> str:
    if ".." in file_path:
        raise ValueError(f"Path traversal detectado: {file_path}")
    resolved = os.path.abspath(file_path)
    enforce_allowed_roots(resolved)
    if must_exist and not os.path.exists(resolved):
        raise ValueError(f"Arquivo nao encontrado: {file_path}")
    if must_exist:
        if os.path.isfile(resolved):
            sz = os.path.getsize(resolved)
            if sz > MAX_FILE_SIZE_BYTES:
                raise ValueError(f"Arquivo excede limite de {MAX_FILE_SIZE_BYTES / (1024*1024)}MB: {sz / (1024*1024):.1f}MB")
    return resolved

def validate_write_path(file_path: str) -> str:
    resolved = validate_path(file_path, False)
    norm = os.path.normpath(resolved)
    for b in get_blocked_roots():
        nb = os.path.normpath(b)
        if norm == nb or norm.startswith(nb + os.sep):
            raise PermissionError(f"Escrita bloqueada em diretorio protegido: {b}")
    return resolved

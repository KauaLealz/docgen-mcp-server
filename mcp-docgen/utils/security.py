import os
import platform
from pathlib import Path

MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024  # 50MB

BLOCKED_DIR_NAMES = {
    ".ssh", ".aws", ".gnupg", ".credentials", ".kube",
    ".docker", ".azure", ".config/gcloud",
}


def _get_blocked_dirs() -> list[Path]:
    blocked: list[Path] = []

    if platform.system() == "Windows":
        sys_root = os.environ.get("SystemRoot", r"C:\Windows")
        blocked.append(Path(sys_root))
        blocked.append(Path("C:/Program Files"))
        blocked.append(Path("C:/Program Files (x86)"))
    else:
        blocked.extend([
            Path("/bin"), Path("/sbin"), Path("/usr"),
            Path("/etc"), Path("/boot"), Path("/sys"), Path("/proc"),
        ])

    home = Path.home()
    for name in BLOCKED_DIR_NAMES:
        blocked.append(home / name)

    return [p.resolve() for p in blocked]


def validate_path(file_path: str, must_exist: bool = False) -> Path:
    """Resolve and validate a file path.

    Blocks path traversal, symlinks, and oversized files.
    """
    resolved = Path(file_path).resolve()

    if ".." in Path(file_path).parts:
        raise ValueError(f"Path traversal detectado: {file_path}")

    if resolved.is_symlink():
        link_target = resolved.resolve(strict=True)
        if link_target != resolved:
            raise ValueError(f"Symlink não permitido: {file_path}")

    if must_exist:
        if not resolved.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
        if resolved.stat().st_size > MAX_FILE_SIZE_BYTES:
            size_mb = resolved.stat().st_size / (1024 * 1024)
            raise ValueError(
                f"Arquivo excede limite de {MAX_FILE_SIZE_BYTES // (1024*1024)}MB: "
                f"{size_mb:.1f}MB"
            )

    return resolved


def validate_write_path(file_path: str) -> Path:
    """Validate a path for write operations.

    Runs base validation (traversal, symlink) then checks against blocked dirs.
    """
    resolved = validate_path(file_path)

    blocked = _get_blocked_dirs()
    resolved_str = str(resolved)
    for b in blocked:
        if resolved_str.startswith(str(b)):
            raise ValueError(
                f"Escrita bloqueada em diretório protegido: {b}"
            )

    return resolved


def validate_image_path(image_path: str) -> Path:
    """Validate an image path exists and has a supported extension."""
    resolved = validate_path(image_path, must_exist=True)
    valid_extensions = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".webp"}
    if resolved.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Formato de imagem não suportado: {resolved.suffix}. "
            f"Suportados: {valid_extensions}"
        )
    return resolved

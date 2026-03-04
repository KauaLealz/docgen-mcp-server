from datetime import datetime
from pathlib import Path

from utils.security import validate_write_path


def generate_output_path(output_path: str) -> Path:
    """Resolve, validate and prepare the output path for writing.

    Creates parent directories if needed. Validates against blocked dirs.
    Existing files at the path will be overwritten.
    """
    if not output_path:
        raise ValueError("output_path é obrigatório - informe o caminho completo do arquivo de saída")
    resolved = validate_write_path(output_path)
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def list_files(directory: str, extension_filter: str | None = None) -> list[dict]:
    """List files in a given directory with metadata."""
    dir_path = Path(directory).resolve()
    if not dir_path.is_dir():
        raise ValueError(f"Diretório não encontrado: {directory}")

    results = []
    for f in sorted(dir_path.iterdir()):
        if not f.is_file():
            continue
        if extension_filter:
            ext = extension_filter if extension_filter.startswith(".") else f".{extension_filter}"
            if f.suffix.lower() != ext.lower():
                continue
        stat = f.stat()
        results.append({
            "name": f.name,
            "path": str(f),
            "size_bytes": stat.st_size,
            "size_human": _human_size(stat.st_size),
            "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
        })

    return results


def _human_size(size_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"

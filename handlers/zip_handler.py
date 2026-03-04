import zipfile
from pathlib import Path

from utils.security import validate_path
from utils.file_utils import generate_output_path


def create_zip(
    file_paths: list[str],
    output_path: str = "",
    compression: str = "deflated",
) -> str:
    """Create a ZIP archive from a list of files.

    Args:
        file_paths: List of absolute paths to include in the archive.
        output_path: Absolute path for the .zip file.
        compression: "deflated" (default, smaller) or "stored" (no compression, faster).

    Returns the absolute path of the generated .zip file.
    """
    if not file_paths:
        raise ValueError("file_paths não pode ser vazio")

    comp = zipfile.ZIP_DEFLATED if compression != "stored" else zipfile.ZIP_STORED

    dest = generate_output_path(output_path)

    with zipfile.ZipFile(str(dest), "w", compression=comp) as zf:
        for fp in file_paths:
            resolved = validate_path(fp, must_exist=True)
            if resolved.is_dir():
                for child in sorted(resolved.rglob("*")):
                    if child.is_file():
                        arcname = str(child.relative_to(resolved.parent))
                        zf.write(str(child), arcname)
            else:
                zf.write(str(resolved), resolved.name)

    return str(dest)


def read_zip(file_path: str) -> dict:
    """List contents of a ZIP archive without extracting.

    Args:
        file_path: Absolute path to the .zip file.

    Returns:
        {
            "entries": [{"name": str, "size_bytes": int, "compressed_bytes": int}],
            "total_files": int,
            "total_size_bytes": int,
            "path": str
        }
    """
    resolved = validate_path(file_path, must_exist=True)

    entries = []
    total_size = 0
    with zipfile.ZipFile(str(resolved), "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            entries.append({
                "name": info.filename,
                "size_bytes": info.file_size,
                "compressed_bytes": info.compress_size,
            })
            total_size += info.file_size

    return {
        "entries": entries,
        "total_files": len(entries),
        "total_size_bytes": total_size,
        "path": str(resolved),
    }

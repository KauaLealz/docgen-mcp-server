from pathlib import Path

from utils.security import validate_path, validate_write_path
from utils.file_utils import generate_output_path


def create_txt(
    content: str,
    output_path: str = "",
    encoding: str = "utf-8",
) -> str:
    """Create a plain text (.txt) file.

    Args:
        content: Text content to write.
        output_path: Absolute path for the output file.
        encoding: File encoding (default utf-8).

    Returns the absolute path of the generated file.
    """
    dest = generate_output_path(output_path)
    dest.write_text(content, encoding=encoding)
    return str(dest)


def read_txt(file_path: str, encoding: str = "utf-8") -> dict:
    """Read a plain text file and return its content with metadata.

    Args:
        file_path: Absolute path to the .txt file.
        encoding: File encoding (default utf-8).

    Returns:
        {
            "content": str,
            "lines": int,
            "size_bytes": int,
            "path": str
        }
    """
    resolved = validate_path(file_path, must_exist=True)
    text = resolved.read_text(encoding=encoding)
    stat = resolved.stat()
    return {
        "content": text,
        "lines": text.count("\n") + (1 if text else 0),
        "size_bytes": stat.st_size,
        "path": str(resolved),
    }


def append_txt(
    file_path: str,
    content: str,
    encoding: str = "utf-8",
) -> str:
    """Append content to an existing text file. Creates the file if it doesn't exist.

    Args:
        file_path: Absolute path to the .txt file.
        content: Text to append.
        encoding: File encoding (default utf-8).

    Returns the absolute path of the file.
    """
    resolved = validate_write_path(file_path)
    resolved.parent.mkdir(parents=True, exist_ok=True)
    with open(resolved, "a", encoding=encoding) as f:
        f.write(content)
    return str(resolved)

"""Allow running as `python -m server` from the project root."""
from server import mcp

if __name__ == "__main__":
    mcp.run(transport="stdio")

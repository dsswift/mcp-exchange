"""Exchange Online MCP Server - Model Context Protocol integration for Microsoft Exchange."""

__version__ = "0.1.0"

from .server import mcp, run_server


def main() -> None:
    """Entry point for the MCP server."""
    run_server()


__all__ = ["main", "mcp", "run_server", "__version__"]

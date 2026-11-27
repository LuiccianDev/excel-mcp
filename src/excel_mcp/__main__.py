"""
Entry point for running excel_mcp as a module.

Starts the Excel MCP server over stdio when invoked via `uvx luiccian-excel-mcp`,
`python -m excel_mcp`, or `uv run excel-mcp`. Runtime settings are read from
environment variables (`EXCEL_MCP_DIRS`, `EXCEL_MCP_TOOLS`, `LOG_LEVEL`,
`EXCEL_MCP_METRICS_LOG`).
"""

import logging
import os
import sys

from excel_mcp.config import (
    ConfigurationError,
    config_manager,
    validate_directory_path,
)
from excel_mcp.server import configure_logging, run_server, run_stdio
from excel_mcp.utils.paths import resolve_allowed_directories

logger = logging.getLogger(__name__)


def validate_and_apply_configuration() -> None:
    """Validate and apply configuration from environment variables."""
    try:
        overrides: dict[str, str] = {}
        if os.environ.get("EXCEL_MCP_DIRS", "").strip():
            allowed = resolve_allowed_directories()
            if allowed:
                for candidate in allowed:
                    validate_directory_path(candidate)
                overrides["directory"] = allowed[0]

        config_manager.reload_configuration(**overrides)
        logger.info("Configuration validated and applied successfully")

    except ConfigurationError as e:
        print(f"Configuration error: {e}", file=sys.stderr)
        logger.error(f"Configuration validation failed: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected configuration error: {e}", file=sys.stderr)
        logger.error(f"Unexpected configuration error: {e}")
        sys.exit(1)


def main() -> None:
    """Start the Excel MCP server on stdio using environment configuration."""
    configure_logging()
    logger.info("MCP Excel Office server starting...")

    if sys.argv[1:]:
        print(
            "excel-mcp takes no arguments; configure it with EXCEL_MCP_DIRS and "
            "EXCEL_MCP_TOOLS environment variables.",
            file=sys.stderr,
        )
        sys.exit(2)

    try:
        validate_and_apply_configuration()
        config = config_manager.config
        logger.info("Server Name: MCP Excel Office Server")
        logger.info(f"File operations directory: {config.file_config.directory}")
        logger.info("Server ready. Listening on stdio transport...")

        server = run_server()
        run_stdio(server)

    except KeyboardInterrupt:
        print("\nServer stopped by user.", file=sys.stderr)
        logger.info("Server stopped by user interrupt")
        sys.exit(0)
    except ConfigurationError as config_error:
        print(f"Configuration error: {config_error}", file=sys.stderr)
        logger.error(f"MCP server configuration failed: {config_error}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error starting server: {e}", file=sys.stderr)
        logger.error(f"Unexpected server startup error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

"""
Entry point for running excel_mcp as a module.

This module provides the command-line interface for the Excel MCP server,
supporting both MCP server mode and standalone CLI operations. It handles
environment variable configuration, command-line argument parsing, and provides
clear error messages for configuration issues.

The module supports execution as a single MCP server command:
- Default MCP Server: `uvx luiccian-excel-mcp`, `python -m excel_mcp`, or `uv run excel-mcp`
- Optional CLI flags for clients without `env` support: `--path`, `--tool-mode`
"""

import logging
import os
import sys
from typing import Annotated

import typer

from excel_mcp.config import (
    ConfigurationError,
    config_manager,
    validate_directory_path,
)
from excel_mcp.server import configure_logging, run_server, run_stdio
from excel_mcp.tools.register_tools import normalize_tool_mode
from excel_mcp.utils.paths import resolve_allowed_directories

logger = logging.getLogger(__name__)


app = typer.Typer(
    name="excel_mcp",
    help="Excel MCP Server - Excel manipulation through Model Context Protocol",
)


def validate_and_apply_configuration(path: str | None = None) -> None:
    """Validate and apply configuration from CLI flags and environment variables."""
    try:
        overrides: dict[str, str] = {}
        if path is not None:
            overrides["directory"] = validate_directory_path(
                path, reject_placeholders=True
            )
        elif os.environ.get("EXCEL_MCP_DIRS", "").strip():
            allowed = resolve_allowed_directories()
            if allowed:
                for candidate in allowed:
                    validate_directory_path(candidate)
                overrides["directory"] = allowed[0]

        config_manager.reload_configuration(**overrides)
        if path is not None:
            os.environ["EXCEL_MCP_DIRS"] = config_manager.config.file_config.directory

        logger.info("Configuration validated and applied successfully")

    except ConfigurationError as e:
        typer.echo(f"Configuration error: {e}", err=True)
        logger.error(f"Configuration validation failed: {e}")
        raise typer.Exit(1) from e
    except Exception as e:
        typer.echo(f"Unexpected configuration error: {e}", err=True)
        logger.error(f"Unexpected configuration error: {e}")
        raise typer.Exit(1) from e


def _start_server(path: str | None = None, tool_mode: str | None = None) -> None:
    """Validate configuration and start the MCP server on stdio."""
    try:
        logger.info("Initializing MCP Excel Office Server...")

        validate_and_apply_configuration(path)
        normalized_mode: str | None = None
        if tool_mode is not None:
            try:
                normalized_mode = normalize_tool_mode(tool_mode)
            except ValueError as exc:
                typer.echo(str(exc), err=True)
                raise typer.Exit(1) from exc

        config = config_manager.config
        logger.info("Starting MCP Excel Office Server...")
        logger.info("Server Name: MCP Excel Office Server")
        logger.info(f"File operations directory: {config.file_config.directory}")
        logger.info("Server ready. Listening on stdio transport...")

        server = run_server(mode=normalized_mode)
        run_stdio(server)

    except KeyboardInterrupt:
        typer.echo("\nServer stopped by user.", err=True)
        logger.info("Server stopped by user interrupt")
        sys.exit(0)
    except typer.Exit:
        raise
    except ConfigurationError as config_error:
        typer.echo(f"Configuration error: {config_error}", err=True)
        logger.error(f"MCP server configuration failed: {config_error}")
        raise typer.Exit(1) from config_error
    except Exception as e:
        typer.echo(f"Unexpected error starting server: {e}", err=True)
        logger.error(f"Unexpected server startup error: {e}")
        raise typer.Exit(1) from e


PathOption = Annotated[
    str | None,
    typer.Option(
        "--path",
        help="Base directory for file operations (overrides EXCEL_MCP_DIRS env var)",
    ),
]
ToolModeOption = Annotated[
    str | None,
    typer.Option(
        "--tool-mode",
        help="Tool registration mode: 'compound' (default, 8 compact tools) or 'full' (33 tools)",
    ),
]


@app.command()
def run(
    path: PathOption = None,
    tool_mode: ToolModeOption = None,
) -> None:
    """Validate configuration and start the MCP server on stdio."""
    _start_server(path=path, tool_mode=tool_mode)


def main() -> None:
    """
    Main entry point for the excel_mcp module.

    This function serves as the primary entry point when the module is executed
    with `python -m excel_mcp` or through the `excel-mcp` CLI script.

    It starts the MCP server on stdio, accepting optional `--path` and
    `--tool-mode` flags when environment variables cannot be passed directly.
    """
    try:
        configure_logging()
        logger.info("MCP Excel Office CLI starting...")
        logger.info(f"Processing command: {' '.join(sys.argv[1:])}")
        app()

    except SystemExit as exc:
        if exc.code not in (None, 0):
            raise
    except KeyboardInterrupt:
        typer.echo("\nOperation cancelled by user.", err=True)
        logger.info("Operation cancelled by user interrupt")
        sys.exit(0)
    except Exception as e:
        typer.echo(f"Unexpected error in MCP Excel CLI: {e}", err=True)
        logger.error(f"Unexpected CLI error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

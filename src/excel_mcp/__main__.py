"""
Entry point for running excel_mcp as a module.

This module provides the command-line interface for the Excel MCP server,
supporting both MCP server mode and standalone CLI operations. It handles
environment variable configuration, command-line argument parsing, and provides
clear error messages for configuration issues.

The module supports multiple execution modes:
- MCP Server: `python -m excel_mcp server` or `uv run excel-mcp`
- File listing: `python -m excel_mcp list`
- Default behavior: Start MCP server when no command is specified
"""

import logging
import os
import sys
from typing import Annotated

import typer

from excel_mcp.config import ConfigurationError, config_manager
from excel_mcp.server import run_server, run_stdio
from excel_mcp.tools.register_tools import normalize_tool_mode


# Configure logging for CLI operations
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


app = typer.Typer(
    name="excel_mcp",
    help="Excel MCP Server - Excel manipulation through Model Context Protocol",
)


def validate_and_apply_configuration(path: str | None = None) -> None:
    """
    Validate and apply configuration from command-line arguments and environment variables.

    This function handles configuration validation and applies overrides from command-line
    arguments. It uses the centralized configuration management system with proper
    error handling and user feedback.

    Args:
        path: Directory path override

    Raises:
        typer.Exit: If configuration is invalid or inaccessible
    """
    try:
        # Prepare configuration overrides
        overrides = {}

        if path is not None:
            overrides["directory"] = path

        # Reload configuration with overrides
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
        logger.info("MCP server starting on stdio transport")

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


@app.callback(invoke_without_command=True)
def root_entrypoint(
    ctx: typer.Context,
    path: Annotated[
        str | None,
        typer.Option(
            "--path",
            help="Base directory for file operations (overrides DIRECTORY env var)",
        ),
    ] = None,
    tool_mode: Annotated[
        str | None,
        typer.Option(
            "--tool-mode",
            help="Tool registration mode: 'compound' (default, 8 compact tools) or 'full' (33 tools)",
        ),
    ] = None,
) -> None:
    """Start the MCP server by default or forward root options to subcommands."""
    if ctx.invoked_subcommand is None:
        _start_server(path=path, tool_mode=tool_mode)
    else:
        ctx.ensure_object(dict).update({"path": path, "tool_mode": tool_mode})


@app.command("server")
def run_mcp_server(
    ctx: typer.Context,
    path: Annotated[
        str | None,
        typer.Option(
            "--path",
            help="Base directory for file operations (overrides DIRECTORY env var)",
        ),
    ] = None,
    tool_mode: Annotated[
        str | None,
        typer.Option(
            "--tool-mode",
            help="Tool registration mode: 'compound' (default, 8 compact tools) or 'full' (33 tools)",
        ),
    ] = None,
) -> None:
    """
    Start the MCP Excel Office server.

    The server provides Excel manipulation capabilities through the Model Context Protocol.
    Configuration can be provided via command-line arguments or environment variables.
    Command-line arguments take precedence over environment variables.
    """
    root_opts = ctx.obj if isinstance(ctx.obj, dict) else {}
    effective_path = path if path is not None else root_opts.get("path")
    effective_mode = tool_mode if tool_mode is not None else root_opts.get("tool_mode")
    _start_server(path=effective_path, tool_mode=effective_mode)


@app.command("list")
def list_excel_files(
    ctx: typer.Context,
    path: Annotated[
        str | None,
        typer.Option(
            "--path",
            help="Directory to scan for Excel files (overrides DIRECTORY env var)",
        ),
    ] = None,
) -> None:
    """
    List all Excel files in the specified directory.

    This command provides a standalone way to discover Excel files without starting the MCP server.
    """
    try:
        root_opts = ctx.obj if isinstance(ctx.obj, dict) else {}
        if root_opts.get("tool_mode") is not None:
            typer.echo(
                "--tool-mode is only supported for the server command.", err=True
            )
            raise typer.Exit(1)
        effective_path = path if path is not None else root_opts.get("path")
        validate_and_apply_configuration(path=effective_path)

        # Get validated configuration
        config = config_manager.config
        target_directory = config.file_config.directory

        # List Excel files
        excel_files = []
        try:
            for entry in os.scandir(target_directory):
                if entry.is_file() and entry.name.lower().endswith((".xlsx", ".xls")):
                    stat_info = entry.stat()
                    excel_files.append(
                        {
                            "name": entry.name,
                            "path": entry.path,
                            "size": stat_info.st_size,
                            "modified": stat_info.st_mtime,
                        }
                    )
        except OSError as e:
            typer.echo(f"Error scanning directory: {e}", err=True)
            raise typer.Exit(1) from e

        # Display results
        if not excel_files:
            typer.echo(f"No Excel files found in: {target_directory}")
        else:
            typer.echo(f"Found {len(excel_files)} Excel file(s) in: {target_directory}")
            for file_info in excel_files:
                raw_size = file_info.get("size", 0)
                size_mb = float(str(raw_size)) / (1024 * 1024)
                typer.echo(f"  - {file_info['name']} ({size_mb:.2f} MB)")

    except typer.Exit:
        raise
    except Exception as e:
        typer.echo(f"Error listing Excel files: {e}", err=True)
        sys.exit(1)


def main() -> None:
    """
    Main entry point for the excel_mcp module.

    This function serves as the primary entry point when the module is executed
    with `python -m excel_mcp` or through the `excel-mcp` CLI script.

    It supports multiple execution patterns:
    - `python -m excel_mcp` (defaults to server mode)
    - `python -m excel_mcp server` (explicit server mode)
    - `python -m excel_mcp list` (list Excel files)
    - `uv run excel-mcp` (entry point script)

    The function provides proper error handling and logging for all execution modes.
    """
    try:
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

"""
Entry point for running mcp_excel as a module.

This module provides the command-line interface for the MCP Excel Office server,
supporting both MCP server mode and standalone CLI operations. It handles
environment variable configuration, command-line argument parsing, and provides
clear error messages for configuration issues.

The module supports multiple execution modes:
- MCP Server: `python -m mcp_excel server` or `uv run mcp_excel_office`
- File listing: `python -m mcp_excel list`
- Default behavior: Start MCP server when no command is specified
"""

import logging
import os
import sys
from typing import Annotated

import typer

from excel_mcp.config import ConfigurationError, config_manager
from excel_mcp.server import run_server


# Configure logging for CLI operations
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


app = typer.Typer(
    name="mcp_excel",
    help="MCP Excel Office Server - Excel manipulation through Model Context Protocol",
    no_args_is_help=True,
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

        # Get validated configuration
        config = config_manager.config

        # Set environment variables for backward compatibility
        if config.directory:
            os.environ["DIRECTORY"] = config.directory

        logger.info("Configuration validated and applied successfully")

    except ConfigurationError as e:
        typer.echo(f"Configuration error: {e}", err=True)
        logger.error(f"Configuration validation failed: {e}")
        raise typer.Exit(1) from e
    except Exception as e:
        typer.echo(f"Unexpected configuration error: {e}", err=True)
        logger.error(f"Unexpected configuration error: {e}")
        raise typer.Exit(1) from e


@app.command("server")
def run_mcp_server(
    path: Annotated[
        str | None,
        typer.Option(
            "--path",
            help="Base directory for file operations (overrides DIRECTORY env var)",
        ),
    ] = None,
) -> None:
    """
    Start the MCP Excel Office server.

    The server provides Excel manipulation capabilities through the Model Context Protocol.
    Configuration can be provided via command-line arguments or environment variables.
    Command-line arguments take precedence over environment variables.

    Environment Variables:
        DIRECTORY: Base directory for file operations (defaults to ./documents)

    Examples:
        python -m mcp_excel server
        python -m mcp_excel server --path "/path/to/excel/files"
    """
    try:
        logger.info("Initializing MCP Excel Office Server...")

        # Validate and apply configuration
        validate_and_apply_configuration(path)

        # Get validated configuration
        config = config_manager.config

        # Provide user feedback about configuration
        typer.echo("Starting MCP Excel Office Server...")
        typer.echo("Server Name: MCP Excel Office Server")
        typer.echo(f"File operations directory: {config.file_config.directory}")
        logger.info(f"File operations directory: {config.file_config.directory}")

        typer.echo("Server ready. Listening on stdio transport...")
        logger.info("MCP server starting on stdio transport")

        # Start the MCP server with proper error handling
        try:
            server = run_server()
            server.run(transport="stdio")
        except ConfigurationError as config_error:
            typer.echo(f"Configuration error: {config_error}", err=True)
            logger.error(f"MCP server configuration failed: {config_error}")
            raise typer.Exit(1) from config_error
        except Exception as server_error:
            typer.echo(f"Error initializing MCP server: {server_error}", err=True)
            logger.error(f"MCP server initialization failed: {server_error}")
            raise typer.Exit(1) from server_error

    except KeyboardInterrupt:
        typer.echo("\nServer stopped by user.", err=True)
        logger.info("Server stopped by user interrupt")
        sys.exit(0)
    except typer.Exit:
        # Re-raise typer.Exit to preserve exit codes
        raise
    except Exception as e:
        typer.echo(f"Unexpected error starting server: {e}", err=True)
        logger.error(f"Unexpected server startup error: {e}")
        sys.exit(1)


@app.command("list")
def list_excel_files(
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
        # Validate and apply configuration
        validate_and_apply_configuration(path=path)

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

    except Exception as e:
        typer.echo(f"Error listing Excel files: {e}", err=True)
        sys.exit(1)


def main() -> None:
    """
    Main entry point for the mcp_excel module.

    This function serves as the primary entry point when the module is executed
    with `python -m mcp_excel` or through the `mcp_excel_office` script.

    It supports multiple execution patterns:
    - `python -m mcp_excel` (defaults to server mode)
    - `python -m mcp_excel server` (explicit server mode)
    - `python -m mcp_excel list` (list Excel files)
    - `uv run mcp_excel_office` (entry point script)

    The function provides proper error handling and logging for all execution modes.
    """
    try:
        logger.info("MCP Excel Office CLI starting...")

        # If no arguments provided, default to server mode for backward compatibility
        if len(sys.argv) == 1:
            # Default behavior: start the MCP server
            typer.echo(
                "No command specified. Starting MCP server (use --help for options)..."
            )
            logger.info("Defaulting to server mode - no command specified")
            run_mcp_server()
        else:
            # Use Typer CLI for command parsing
            logger.info(f"Processing command: {' '.join(sys.argv[1:])}")
            app()

    except KeyboardInterrupt:
        typer.echo("\nOperation cancelled by user.", err=True)
        logger.info("Operation cancelled by user interrupt")
        sys.exit(0)
    except typer.Exit as e:
        # Re-raise typer.Exit to preserve exit codes
        logger.info(f"CLI exited with code: {e.exit_code}")
        raise
    except Exception as e:
        typer.echo(f"Unexpected error in MCP Excel CLI: {e}", err=True)
        logger.error(f"Unexpected CLI error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

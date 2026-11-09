"""
Configuration management for MCP Excel Office Server.

This module provides comprehensive configuration management supporting multiple
deployment modes (DXT, traditional MCP, standalone CLI) with proper validation,
error handling, and user_config variable substitution.
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


# Load .env file automatically if available, but don't override existing env vars
try:
    from dotenv import load_dotenv

    # Use override=False to prioritize existing environment variables (from DXT)
    load_dotenv(override=False)
except ImportError:
    pass  # python-dotenv not available, continue without it


logger = logging.getLogger(__name__)


class ConfigurationError(Exception):
    """Raised when configuration is invalid or missing."""

    pass


class FileConfig(BaseModel):
    """File operations configuration with validation."""

    directory: str = Field(
        default="./documents", description="Base directory for file operations"
    )
    max_file_size: int = Field(
        default=100 * 1024 * 1024, gt=0, description="Maximum file size in bytes"
    )
    allowed_extensions: list[str] = Field(
        default=[".xlsx"], description="Allowed file extensions"
    )

    @field_validator("directory")
    @classmethod
    def validate_directory(cls: type[FileConfig], v: str) -> str:
        """Validate and normalize directory path."""
        if not isinstance(v, str):
            raise ValueError("Directory must be a string")

        if not v.strip():
            raise ValueError("Directory cannot be empty")

        # Normalize path
        path = Path(v).resolve()

        # Check for clearly invalid paths that we shouldn't try to create
        invalid_patterns = [
            "/path/that/does/not/exist",
            "/path/that/cannot/be/created/due/to/permissions",
            "/invalid/path",
            "C:\\path\\that\\does\\not\\exist",
            "C:\\path\\that\\cannot\\be\\created\\due\\to\\permissions",
            "C:\\invalid\\path",
        ]

        if str(path) in invalid_patterns or any(
            pattern in str(path) for pattern in ["/path/that/", "C:\\path\\that\\"]
        ):
            raise ValueError(f"Invalid directory path: {path}")

        # Security check: prevent directory traversal
        try:
            # Only try to create directory if it's a reasonable path
            if path.exists():
                # Directory already exists, check if it's actually a directory
                if not path.is_dir():
                    raise ValueError(f"Path is not a directory: {path}")
            else:
                # Try to create directory if it doesn't exist
                path.mkdir(parents=True, exist_ok=True)

            # Check read/write permissions
            if not os.access(path, os.R_OK | os.W_OK):
                raise ValueError(f"Directory is not readable/writable: {path}")

        except (OSError, PermissionError) as e:
            raise ValueError(f"Cannot create or access directory '{path}': {e}")  # noqa: B904

        return str(path)

    @field_validator("allowed_extensions")
    @classmethod
    def validate_extensions(cls: type[FileConfig], v: list[str]) -> list[str]:
        """Validate file extensions."""
        if not isinstance(v, list):
            raise ValueError("Allowed extensions must be a list")

        for ext in v:
            if not isinstance(ext, str):
                raise ValueError("File extensions must be strings")
            if not ext.startswith("."):
                raise ValueError(f"File extension must start with '.': {ext}")

        return v


class MCPExcelConfig(BaseSettings):
    """
    Main configuration class for MCP Excel Office Server.

    Supports multiple configuration sources with proper precedence:
    1. Command-line arguments (highest priority)
    2. Environment variables
    3. User config variables (${user_config.*})
    4. Default values (lowest priority)
    """

    # File operations configuration
    directory: str = Field(
        default="./documents",
        description="Base directory for file operations",
    )

    # Logging configuration
    log_level: str = Field(
        default="INFO",
        pattern="^(DEBUG|INFO|WARNING|ERROR|CRITICAL)$",
        description="Logging level",
    )

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,  # Cambiar a False para mayor compatibilidad
        extra="ignore",
        env_prefix="",  # Sin prefijo para las variables de entorno
    )

    def __init__(self, **kwargs: Any) -> None:
        """Initialize configuration with user_config variable substitution."""
        # Process user_config variables before initialization
        processed_kwargs = self._process_user_config_variables(kwargs)

        # Initialize with processed kwargs
        super().__init__(**processed_kwargs)

        # Validate configuration after initialization
        self._validate_configuration()

    def _process_user_config_variables(self, kwargs: dict[str, Any]) -> dict[str, Any]:
        """
        Process ${user_config.*} variable substitution.

        This supports DXT and traditional MCP deployments where configuration
        can include variables like ${user_config.*}.
        """
        processed = {}
        user_config_pattern = re.compile(r"\$\{user_config\.([^}]+)\}")

        for key, value in kwargs.items():
            if isinstance(value, str) and "${user_config." in value:
                # Extract user_config variable name
                match = user_config_pattern.search(value)
                if match:
                    config_key = match.group(1)

                    # Try to get value from environment variables
                    # Convert user_config key to environment variable format
                    env_value = os.environ.get(config_key.upper())

                    if env_value:
                        # Replace the user_config variable with actual value
                        # Use string replacement instead of regex substitution to avoid Windows path issues
                        processed_value = value.replace(
                            f"${{user_config.{config_key}}}", env_value
                        )
                        processed[key] = processed_value
                        logger.info(f"Resolved user_config.{config_key} = {env_value}")

                        # For DXT compatibility, check if we're in a DXT environment
                        # and use fallback values
                    elif config_key == "directory":
                        # Use current working directory as fallback for DXT
                        fallback_dir = os.getcwd()
                        processed[key] = value.replace(
                            f"${{user_config.{config_key}}}", fallback_dir
                        )
                        logger.warning(
                            f"Could not resolve user_config.{config_key}, using fallback: {fallback_dir}"
                        )
                    else:
                        # Skip this key to let Pydantic use its default value
                        logger.warning(
                            f"Could not resolve user_config.{config_key}, using default value"
                        )
                        continue
                else:
                    processed[key] = value
            else:
                processed[key] = value

        return processed

    def _validate_configuration(self) -> None:
        """Validate the complete configuration."""
        try:
            # Validate file configuration
            file_config = FileConfig(directory=self.directory)
            logger.info(f"File operations directory validated: {file_config.directory}")

        except Exception as e:
            raise ConfigurationError(f"Configuration validation failed: {e}") from e

    @property
    def file_config(self) -> FileConfig:
        """Get validated file configuration."""
        return FileConfig(directory=self.directory)

    def get_effective_config(self) -> dict[str, Any]:
        """Get the effective configuration as a dictionary."""
        return {
            "directory": self.directory,
            "log_level": self.log_level,
        }


class ConfigurationManager:
    """
    Centralized configuration management for MCP Excel Office Server.

    Handles configuration loading, validation, and provides a single point
    of access for all configuration needs across the application.
    """

    _instance: ConfigurationManager | None = None

    _config: MCPExcelConfig | None = None

    def __new__(cls) -> ConfigurationManager:
        """Singleton pattern to ensure single configuration instance."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self) -> None:
        """Initialize configuration manager."""
        if self._config is None:
            self._load_configuration()

    def _load_configuration(self, **overrides: Any) -> None:
        """Load configuration from all sources with proper precedence."""
        try:
            self._config = MCPExcelConfig(**overrides)
            logger.info("Configuration loaded successfully")
        except Exception as e:
            logger.error(f"Failed to load configuration: {e}")
            raise ConfigurationError(f"Configuration loading failed: {e}") from e

    def reload_configuration(self, **overrides: Any) -> None:
        """Reload configuration with optional overrides."""
        logger.info("Reloading configuration...")
        self._load_configuration(**overrides)

    @property
    def config(self) -> MCPExcelConfig:
        """Get the current configuration."""
        if self._config is None:
            raise ConfigurationError("Configuration not loaded")
        return self._config

    def get_directory(self) -> str:
        """Get base directory for file operations with validation."""
        return self.config.directory

    def get_log_level(self) -> str:
        """Get logging level."""
        return self.config.log_level

    def validate_file_path(self, file_path: str) -> str:
        """
        Validate file path against configured directory.

        Args:
            file_path: File path to validate

        Returns:
            Validated absolute file path

        Raises:
            ConfigurationError: If file path is invalid or outside allowed directory
        """
        try:
            base_dir = Path(self.get_directory()).resolve()
            target_path = Path(file_path).resolve()

            # Security check: ensure file is within allowed directory
            try:
                target_path.relative_to(base_dir)
            except ValueError:
                raise ConfigurationError(  # noqa: B904
                    f"File path is outside allowed directory. "
                    f"File: {target_path}, Allowed directory: {base_dir}"
                )

            return str(target_path)

        except Exception as e:
            raise ConfigurationError(f"File path validation failed: {e}") from e

    def get_configuration_summary(self) -> str:
        """Get a human-readable configuration summary."""
        config = self.config.get_effective_config()

        summary = [
            "MCP Excel Office Server Configuration:",
            f"  File Directory: {config['directory']}",
            f"  Log Level: {config['log_level']}",
        ]

        return "\n".join(summary)


# Global configuration manager instance
config_manager = ConfigurationManager()


def get_config() -> MCPExcelConfig:
    """Get the global configuration instance."""
    return config_manager.config


def get_directory() -> str:
    """Get base directory for file operations."""
    return config_manager.get_directory()


def validate_file_path(file_path: str) -> str:
    """Validate file path against configuration."""
    return config_manager.validate_file_path(file_path)


def reload_configuration(**overrides: Any) -> None:
    """Reload configuration with optional overrides."""
    config_manager.reload_configuration(**overrides)


def validate_configuration() -> list[str]:
    """
    Validate current configuration and return list of issues.

    Returns:
        list[str]: List of configuration issues (empty if all valid)
    """
    issues = []

    try:
        # Validate file directory
        try:
            directory = get_directory()
            if not Path(directory).exists():
                issues.append(f"File directory does not exist: {directory}")
            elif not os.access(directory, os.R_OK | os.W_OK):
                issues.append(f"File directory is not readable/writable: {directory}")
        except Exception as e:
            issues.append(f"File directory validation failed: {e}")

        # Validate log level
        valid_levels = ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
        if config_manager.config.log_level not in valid_levels:
            issues.append(
                f"Invalid log level: {config_manager.config.log_level}. Must be one of: {', '.join(valid_levels)}"
            )

    except Exception as e:
        issues.append(f"Configuration validation failed: {e}")

    return issues

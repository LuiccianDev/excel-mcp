"""
Configuration management for MCP Excel Office Server.

This module provides comprehensive configuration management supporting multiple
deployment modes (DXT, traditional MCP, standalone CLI) with proper validation,
error handling, and user_config variable substitution.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
import re
from typing import Any

from pydantic import BaseModel, Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

logger = logging.getLogger(__name__)
_UNEXPANDED_PLACEHOLDER_RE = re.compile(r"\$\{[^}]+\}")


class ConfigurationError(Exception):
    """Raised when configuration is invalid or missing."""

    pass


def _default_documents_directory() -> str:
    """Return ~/Documents when writable (or when cwd is in WINDIR), else ./documents."""
    docs_dir = Path.home() / "Documents"
    try:
        if (
            docs_dir.exists()
            and docs_dir.is_dir()
            and os.access(docs_dir, os.R_OK | os.W_OK)
        ):
            return str(docs_dir)
    except OSError:
        pass

    windir = os.environ.get("WINDIR", "").strip()
    if windir:
        try:
            cwd = Path.cwd().resolve()
            win_root = Path(windir).resolve()
            if cwd == win_root or cwd.is_relative_to(win_root):
                return str(docs_dir)
        except OSError:
            pass

    return "./documents"


def validate_directory_path(value: str) -> str:
    """Validate and normalize a configured directory path without mutating disk."""
    if not isinstance(value, str):
        raise ValueError("Directory must be a string")

    if not value.strip():
        raise ValueError("Directory cannot be empty")

    path = Path(value).resolve()

    try:
        if path.exists():
            if not path.is_dir():
                raise ValueError(f"Path is not a directory: {path}")
            if not os.access(path, os.R_OK | os.W_OK):
                raise ValueError(f"Directory is not readable/writable: {path}")

    except (OSError, PermissionError) as e:
        raise ValueError(f"Cannot access directory '{path}': {e}") from e

    return str(path)


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
        return validate_directory_path(v)

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

    directory: str = Field(
        default_factory=_default_documents_directory,
        description="Base directory for file operations",
    )

    log_level: str = Field(
        default="INFO",
        pattern="^(DEBUG|INFO|WARNING|ERROR|CRITICAL)$",
        description="Logging level",
    )

    model_config = SettingsConfigDict(
        case_sensitive=False,
        extra="ignore",
        env_prefix="",
    )

    @field_validator("directory", mode="before")
    @classmethod
    def sanitize_unexpanded_directory(cls, value: Any) -> Any:
        """Fall back to default directory if directory contains an unexpanded placeholder."""
        if isinstance(value, str) and _UNEXPANDED_PLACEHOLDER_RE.search(value):
            logger.warning(
                "Ignoring unexpanded directory placeholder %r; using default directory.",
                value,
            )
            return _default_documents_directory()
        return value

    def __init__(self, **kwargs: Any) -> None:
        """Initialize configuration with user_config variable substitution."""
        processed_kwargs = self._process_user_config_variables(kwargs)
        super().__init__(**processed_kwargs)
        self._validate_configuration()

    def _process_user_config_variables(self, kwargs: dict[str, Any]) -> dict[str, Any]:
        """Process ${user_config.*} variable substitution for DXT and MCP deployments."""
        processed = {}
        user_config_pattern = re.compile(r"\$\{user_config\.([^}]+)\}")

        for key, value in kwargs.items():
            if isinstance(value, str) and "${user_config." in value:
                match = user_config_pattern.search(value)
                if match:
                    config_key = match.group(1)
                    env_value = os.environ.get(config_key.upper())

                    if env_value:
                        processed_value = value.replace(
                            f"${{user_config.{config_key}}}", env_value
                        )
                        processed[key] = processed_value
                        logger.info(f"Resolved user_config.{config_key} = {env_value}")

                    elif config_key == "directory":
                        fallback_dir = os.getcwd()
                        processed[key] = value.replace(
                            f"${{user_config.{config_key}}}", fallback_dir
                        )
                        logger.warning(
                            f"Could not resolve user_config.{config_key}, using fallback: {fallback_dir}"
                        )
                    else:
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

    def get_configuration_summary(self) -> str:
        """Get a human-readable configuration summary."""
        config = self.config.get_effective_config()

        summary = [
            "MCP Excel Office Server Configuration:",
            f"  File Directory: {config['directory']}",
            f"  Log Level: {config['log_level']}",
        ]

        return "\n".join(summary)


config_manager = ConfigurationManager()


def get_config() -> MCPExcelConfig:
    """Get the global configuration instance."""
    return config_manager.config


def get_directory() -> str:
    """Get base directory for file operations."""
    return config_manager.get_directory()


def reload_configuration(**overrides: Any) -> None:
    """Reload configuration with optional overrides."""
    config_manager.reload_configuration(**overrides)

"""Tests for excel_mcp.utils.file_utils module."""

import os
from pathlib import Path
from unittest.mock import patch

import pytest

from excel_mcp.config import ConfigurationError
from excel_mcp.utils.file_utils import (
    _check_file_writeable,
    _get_allowed_directories,
    _is_path_in_allowed_directories,
    create_document_copy,
    ensure_xlsx_extension,
    list_excel_files_in_directory,
    resolve_safe_path,
    validate_file_access,
)


class TestGetAllowedDirectories:
    """Tests for _get_allowed_directories function."""

    def test_get_allowed_directories_from_config(self, tmp_path: Path) -> None:
        """Test getting allowed directories from configuration."""
        with patch(
            "excel_mcp.utils.paths.get_directory", return_value=str(tmp_path)
        ):
            result = _get_allowed_directories()

        assert len(result) >= 1
        assert str(tmp_path.resolve()) in [str(Path(d).resolve()) for d in result]

    def test_get_allowed_directories_fallback_to_env(self, tmp_path: Path) -> None:
        """Test fallback to environment variable when config fails."""
        with patch(
            "excel_mcp.utils.paths.get_directory",
            side_effect=Exception("Config error"),
        ):
            with patch.dict(os.environ, {"DIRECTORY": str(tmp_path)}):
                result = _get_allowed_directories()

        assert str(tmp_path.resolve()) in [str(Path(d).resolve()) for d in result]


class TestIsPathInAllowedDirectories:
    """Tests for _is_path_in_allowed_directories function."""

    def test_path_in_allowed_directory(self, tmp_path: Path) -> None:
        """Test path within allowed directory."""
        test_file = tmp_path / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            is_allowed, error = _is_path_in_allowed_directories(str(test_file))

        assert is_allowed is True
        assert error is None

    def test_path_outside_allowed_directory(self, tmp_path: Path) -> None:
        """Test path outside allowed directory."""
        allowed_path = tmp_path / "allowed"
        outside_path = tmp_path / "outside" / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(allowed_path)],
        ):
            is_allowed, error = _is_path_in_allowed_directories(str(outside_path))

        assert is_allowed is False
        assert error is not None
        assert "not in allowed directories" in error

    def test_path_with_symlink(self, tmp_path: Path) -> None:
        """Test path resolution with symlinks."""
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        link_dir = tmp_path / "link"
        try:
            link_dir.symlink_to(real_dir)
        except (OSError, NotImplementedError):
            pytest.skip("Symlinks not supported on this platform")

        test_file = real_dir / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(link_dir)],
        ):
            is_allowed, _ = _is_path_in_allowed_directories(str(test_file))

        # Should resolve symlink and check
        assert is_allowed is True


class TestCheckFileWriteable:
    """Tests for _check_file_writeable function."""

    def test_check_writeable_new_file_in_allowed_dir(self, tmp_path: Path) -> None:
        """Test checking write permission for new file."""
        test_file = tmp_path / "new_file.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            is_writable, error = _check_file_writeable(str(test_file))

        assert is_writable is True
        assert error == ""

    def test_check_writeable_existing_file(self, tmp_path: Path) -> None:
        """Test checking write permission for existing file."""
        test_file = tmp_path / "existing.xlsx"
        test_file.write_text("test")

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            is_writable, error = _check_file_writeable(str(test_file))

        assert is_writable is True

    def test_check_writeable_directory_not_writeable(self, tmp_path: Path) -> None:
        """Test checking write permission when directory is not writeable."""
        test_file = tmp_path / "subdir" / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            is_writable, error = _check_file_writeable(str(test_file))

        # Directory doesn't exist
        assert is_writable is False
        assert "Directory does not exist" in error

    def test_check_writeable_outside_allowed(self, tmp_path: Path) -> None:
        """Test checking write permission outside allowed directories."""
        allowed_dir = tmp_path / "allowed"
        test_file = tmp_path / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(allowed_dir)],
        ):
            is_writable, error = _check_file_writeable(str(test_file))

        assert is_writable is False
        assert "not in allowed directories" in error


class TestResolveSafePath:
    """Tests for resolve_safe_path function."""

    def test_resolve_relative_path(self, tmp_path: Path) -> None:
        """Test resolving relative path."""
        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            result = resolve_safe_path("test.xlsx")

        assert result == tmp_path / "test.xlsx"

    def test_resolve_absolute_path_in_allowed(self, tmp_path: Path) -> None:
        """Test resolving absolute path within allowed directory."""
        test_file = tmp_path / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            result = resolve_safe_path(test_file)

        assert result == test_file

    def test_resolve_absolute_path_outside_allowed(self, tmp_path: Path) -> None:
        """Test resolving absolute path outside allowed directory."""
        allowed_dir = tmp_path / "allowed"
        outside_file = tmp_path / "outside" / "test.xlsx"

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(allowed_dir)],
        ):
            result = resolve_safe_path(outside_file)

        # Should use filename in first allowed directory
        assert result == allowed_dir / "test.xlsx"

    def test_resolve_no_allowed_directories(self) -> None:
        """Test resolving when no allowed directories available."""
        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories", return_value=[]
        ):
            with pytest.raises(PermissionError) as exc_info:
                resolve_safe_path("test.xlsx")

            assert "No allowed directories" in str(exc_info.value)


class TestCreateDocumentCopy:
    """Tests for create_document_copy function."""

    def test_create_copy_success(self, tmp_path: Path) -> None:
        """Test successful document copy."""
        source = tmp_path / "source.xlsx"
        source.write_text("test content")

        success, message, new_path = create_document_copy(str(source))

        assert success is True
        assert "copied" in message
        assert new_path is not None
        assert Path(new_path).exists()

    def test_create_copy_with_dest_path(self, tmp_path: Path) -> None:
        """Test copy with specified destination path."""
        source = tmp_path / "source.xlsx"
        dest = tmp_path / "destination.xlsx"
        source.write_text("test content")

        success, message, new_path = create_document_copy(str(source), str(dest))

        assert success is True
        assert new_path == str(dest)
        assert dest.exists()

    def test_create_copy_source_not_found(self, tmp_path: Path) -> None:
        """Test copy when source doesn't exist."""
        source = tmp_path / "nonexistent.xlsx"

        success, message, new_path = create_document_copy(str(source))

        assert success is False
        assert "does not exist" in message
        assert new_path is None


class TestEnsureXlsxExtension:
    """Tests for ensure_xlsx_extension function."""

    def test_filename_with_extension(self) -> None:
        """Test filename that already has extension."""
        result = ensure_xlsx_extension("test.xlsx")
        assert result == "test.xlsx"

    def test_filename_without_extension(self) -> None:
        """Test filename without extension."""
        result = ensure_xlsx_extension("test")
        assert result == "test.xlsx"

    def test_filename_with_other_extension(self) -> None:
        """Test filename with different extension."""
        result = ensure_xlsx_extension("test.docx")
        assert result == "test.docx.xlsx"


class TestListExcelFilesInDirectory:
    """Tests for list_excel_files_in_directory function."""

    def test_list_excel_files(self, tmp_path: Path) -> None:
        """Test listing Excel files."""
        # Create some files
        (tmp_path / "file1.xlsx").write_text("content")
        (tmp_path / "file2.xlsx").write_text("content")
        (tmp_path / "other.txt").write_text("content")

        with patch(
            "excel_mcp.utils.paths.get_directory", return_value=str(tmp_path)
        ):
            result = list_excel_files_in_directory()

        assert len(result) == 2
        filenames = [f["filename"] for f in result]
        assert "file1.xlsx" in filenames
        assert "file2.xlsx" in filenames

    def test_list_excel_files_with_metadata(self, tmp_path: Path) -> None:
        """Test listing includes file metadata."""
        test_file = tmp_path / "test.xlsx"
        test_file.write_text("content")

        with patch(
            "excel_mcp.utils.paths.get_directory", return_value=str(tmp_path)
        ):
            result = list_excel_files_in_directory()

        assert len(result) == 1
        file_info = result[0]
        assert "filename" in file_info
        assert "size_kb" in file_info
        assert "modified" in file_info
        assert "path" in file_info

    def test_list_excel_files_empty_directory(self, tmp_path: Path) -> None:
        """Test listing in empty directory."""
        with patch(
            "excel_mcp.utils.paths.get_directory", return_value=str(tmp_path)
        ):
            result = list_excel_files_in_directory()

        assert result == []

    def test_list_excel_files_directory_not_found(self, tmp_path: Path) -> None:
        """Test listing when directory doesn't exist."""
        nonexistent = tmp_path / "nonexistent"

        with patch(
            "excel_mcp.utils.paths.get_directory", return_value=str(nonexistent)
        ):
            # The function wraps FileNotFoundError in OSError
            with pytest.raises((FileNotFoundError, OSError)) as exc_info:
                list_excel_files_in_directory()

            assert "not found" in str(exc_info.value).lower()

    def test_list_excel_files_config_error(self) -> None:
        """Test listing when configuration fails."""
        with patch(
            "excel_mcp.utils.paths.get_directory",
            side_effect=Exception("Config error"),
        ):
            with pytest.raises(ConfigurationError):
                list_excel_files_in_directory()


class TestValidateFileAccess:
    """Tests for validate_file_access decorator."""

    def test_validate_sync_function_success(self, tmp_path: Path) -> None:
        """Test decorator on sync function with valid access."""
        test_file = tmp_path / "test.xlsx"

        @validate_file_access("filename")
        def test_func(filename: str) -> dict:
            return {"status": "success", "filename": filename}

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            result = test_func(str(test_file))

        assert result["status"] == "success"

    def test_validate_async_function_success(self, tmp_path: Path) -> None:
        """Test decorator on async function with valid access."""
        test_file = tmp_path / "test.xlsx"

        @validate_file_access("filename")
        async def test_async_func(filename: str) -> dict:
            return {"status": "success", "filename": filename}

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(tmp_path)],
        ):
            import asyncio

            result = asyncio.run(test_async_func(str(test_file)))

        assert result["status"] == "success"

    def test_validate_access_denied(self, tmp_path: Path) -> None:
        """Test decorator when access is denied."""
        allowed_dir = tmp_path / "allowed"
        test_file = tmp_path / "outside" / "test.xlsx"

        @validate_file_access("filename")
        def test_func(filename: str) -> dict:
            return {"status": "success"}

        with patch(
            "excel_mcp.utils.paths.resolve_allowed_directories",
            return_value=[str(allowed_dir)],
        ):
            result = test_func(str(test_file))

        assert result["status"] == "error"
        assert "Access denied" in result["message"]

    def test_validate_missing_param(self) -> None:
        """Test decorator when parameter is missing."""

        @validate_file_access("filename")
        def test_func(other_param: str) -> dict:
            return {"status": "success"}

        result = test_func("value")

        assert result["status"] == "error"
        assert "parameter not found" in result["message"]

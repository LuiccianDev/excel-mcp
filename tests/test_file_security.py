"""
Security-focused test suite for path validation and symlink security.

These tests verify that the path validation subsystem:
1. Resolves symlinks prior to validation (prevents symlink attacks)
2. Blocks path traversal attempts
3. Correctly validates relative and absolute file paths
"""

import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

# Functions under test
from excel_mcp.utils.file_utils import (
    _get_allowed_directories,
    _is_path_in_allowed_directories,
    resolve_safe_path,
)


class TestSymlinkSecurity:
    """Test suite to verify protection against symlink attacks."""

    def test_symlink_outside_allowed_directory_blocked(self):
        """Test that symlinks pointing outside allowed directories are blocked."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with tempfile.TemporaryDirectory() as outside_dir:
                # Create sensitive file outside allowed directory
                outside_file = os.path.join(outside_dir, "secret.txt")
                with open(outside_file, "w") as f:
                    f.write("secret data")

                # Create symlink inside allowed directory pointing outside
                symlink_path = os.path.join(allowed_dir, "malicious_link")

                # Create symlink (only on supported platforms)
                try:
                    os.symlink(outside_file, symlink_path)
                except (OSError, NotImplementedError):
                    pytest.skip("Symlinks not supported on this platform")

                # Mock allowed directory to allowed_dir
                with patch(
                    "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
                ):
                    # Attempt to validate symlink - must fail
                    is_allowed, error = _is_path_in_allowed_directories(symlink_path)

                    assert not is_allowed, (
                        f"Symlink to outside directory should be blocked. Error: {error}"
                    )
                    assert (
                        "not in allowed directories" in error
                        or "resolved" in error.lower()
                    )

    def test_symlink_inside_allowed_directory_allowed(self):
        """Test that symlinks pointing within allowed directories are permitted."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            # Create legitimate file inside allowed directory
            inside_file = os.path.join(allowed_dir, "real_file.txt")
            with open(inside_file, "w") as f:
                f.write("data")

            # Create symlink inside allowed directory pointing inside
            symlink_path = os.path.join(allowed_dir, "link_to_file")

            try:
                os.symlink(inside_file, symlink_path)
            except (OSError, NotImplementedError):
                pytest.skip("Symlinks not supported on this platform")

            # Mock allowed directory to allowed_dir
            with patch(
                "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
            ):
                # Validate symlink - must pass
                is_allowed, error = _is_path_in_allowed_directories(symlink_path)

                assert is_allowed, (
                    f"Symlink inside allowed directory should be permitted. Error: {error}"
                )


class TestPathTraversalSecurity:
    """Test suite to verify protection against path traversal attacks."""

    def test_path_traversal_blocked(self):
        """Test that path traversal attempts are blocked."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            # Create legitimate file
            legit_file = os.path.join(allowed_dir, "file.xlsx")
            with open(legit_file, "w") as f:
                f.write("data")

            with patch(
                "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
            ):
                # Attempt path traversal - must fail
                malicious_path = os.path.join(allowed_dir, "..", "..", "etc", "passwd")
                is_allowed, error = _is_path_in_allowed_directories(malicious_path)

                assert not is_allowed, (
                    f"Path traversal should be blocked: {malicious_path}"
                )

    def test_relative_path_with_dotdot_blocked(self):
        """Test that relative paths with parent directory navigation (..) are blocked."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with patch(
                "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
            ):
                # Attempt relative path traversal
                is_allowed, error = _is_path_in_allowed_directories(
                    "../../../etc/passwd"
                )

                assert not is_allowed, "Relative path traversal should be blocked"

    def test_absolute_path_outside_blocked(self):
        """Test that absolute paths outside allowed directories are blocked."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with tempfile.TemporaryDirectory() as outside_dir:
                outside_file = os.path.join(outside_dir, "file.xlsx")

                with patch(
                    "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
                ):
                    # Attempt to access file outside allowed directory
                    is_allowed, error = _is_path_in_allowed_directories(outside_file)

                    assert not is_allowed, (
                        "Absolute path outside allowed dir should be blocked"
                    )


class TestResolveSafePath:
    """Test suite for resolve_safe_path function."""

    def test_resolve_safe_path_normalizes_relative_path(self):
        """Test that relative paths are normalized into allowed directory."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with patch(
                "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
            ):
                result = resolve_safe_path("document.xlsx")

                assert str(result) == os.path.join(allowed_dir, "document.xlsx")

    def test_resolve_safe_path_blocks_absolute_outside(self):
        """Test that absolute paths outside allowed directory are rewritten safely."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with tempfile.TemporaryDirectory() as outside_dir:
                outside_file = os.path.join(outside_dir, "file.xlsx")

                with patch(
                    "excel_mcp.utils.paths.get_directory", return_value=allowed_dir
                ):
                    result = resolve_safe_path(outside_file)

                    # Should be safely rewritten to allowed directory
                    expected = os.path.join(allowed_dir, "file.xlsx")
                    assert str(result) == expected


class TestAllowedDirectories:
    """Test suite to verify allowed directories configuration."""

    def test_get_allowed_directories_resolves_symlinks(self):
        """Test that allowed directories resolve symlinks cleanly."""
        with tempfile.TemporaryDirectory() as real_dir:
            # Create a symlink to the directory
            symlink_dir = real_dir + "_link"

            try:
                os.symlink(real_dir, symlink_dir)
            except (OSError, NotImplementedError):
                pytest.skip("Symlinks not supported on this platform")

            with patch(
                "excel_mcp.utils.paths.get_directory", return_value=symlink_dir
            ):
                dirs = _get_allowed_directories()

                # Directory must be fully resolved (symlink-free)
                assert len(dirs) == 1
                assert Path(dirs[0]).resolve() == Path(real_dir).resolve()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

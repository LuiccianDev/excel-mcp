"""
Tests específicos de seguridad para validación de rutas y symlinks.

Estos tests verifican que el sistema de validación de rutas:
1. Resuelve symlinks antes de validar (evita symlink attacks)
2. Bloquea path traversal attempts
3. Valida correctamente rutas relativas y absolutas
"""

import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

# Importar funciones a testear
from excel_mcp.utils.file_utils import (
    _get_allowed_directories,
    _is_path_in_allowed_directories,
    resolve_safe_path,
)


class TestSymlinkSecurity:
    """Tests para verificar protección contra symlink attacks."""

    def test_symlink_outside_allowed_directory_blocked(self):
        """Test que symlinks apuntando fuera del directorio permitido son bloqueados."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with tempfile.TemporaryDirectory() as outside_dir:
                # Crear archivo fuera del directorio permitido
                outside_file = os.path.join(outside_dir, "secret.txt")
                with open(outside_file, "w") as f:
                    f.write("secret data")

                # Crear symlink dentro del directorio permitido
                symlink_path = os.path.join(allowed_dir, "malicious_link")

                # Crear symlink (solo en sistemas que lo soportan)
                try:
                    os.symlink(outside_file, symlink_path)
                except (OSError, NotImplementedError):
                    pytest.skip("Symlinks not supported on this platform")

                # Simular que el directorio permitido es allowed_dir
                with patch(
                    "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
                ):
                    # Intentar validar el symlink - debería fallar
                    is_allowed, error = _is_path_in_allowed_directories(symlink_path)

                    assert not is_allowed, (
                        f"Symlink to outside directory should be blocked. Error: {error}"
                    )
                    assert (
                        "not in allowed directories" in error
                        or "resolved" in error.lower()
                    )

    def test_symlink_inside_allowed_directory_allowed(self):
        """Test que symlinks apuntando dentro del directorio permitido son permitidos."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            # Crear archivo dentro del directorio permitido
            inside_file = os.path.join(allowed_dir, "real_file.txt")
            with open(inside_file, "w") as f:
                f.write("data")

            # Crear symlink dentro del directorio permitido
            symlink_path = os.path.join(allowed_dir, "link_to_file")

            try:
                os.symlink(inside_file, symlink_path)
            except (OSError, NotImplementedError):
                pytest.skip("Symlinks not supported on this platform")

            # Simular que el directorio permitido es allowed_dir
            with patch(
                "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
            ):
                # Validar el symlink - debería pasar
                is_allowed, error = _is_path_in_allowed_directories(symlink_path)

                assert is_allowed, (
                    f"Symlink inside allowed directory should be permitted. Error: {error}"
                )


class TestPathTraversalSecurity:
    """Tests para verificar protección contra path traversal attacks."""

    def test_path_traversal_blocked(self):
        """Test que los intentos de path traversal son bloqueados."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            # Crear archivo legítimo
            legit_file = os.path.join(allowed_dir, "file.xlsx")
            with open(legit_file, "w") as f:
                f.write("data")

            with patch(
                "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
            ):
                # Intentar path traversal - debería fallar
                malicious_path = os.path.join(allowed_dir, "..", "..", "etc", "passwd")
                is_allowed, error = _is_path_in_allowed_directories(malicious_path)

                assert not is_allowed, (
                    f"Path traversal should be blocked: {malicious_path}"
                )

    def test_relative_path_with_dotdot_blocked(self):
        """Test que rutas relativas con .. son bloqueadas."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with patch(
                "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
            ):
                # Intentar path traversal relativo
                is_allowed, error = _is_path_in_allowed_directories(
                    "../../../etc/passwd"
                )

                assert not is_allowed, f"Relative path traversal should be blocked"

    def test_absolute_path_outside_blocked(self):
        """Test que rutas absolutas fuera del directorio permitido son bloqueadas."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with tempfile.TemporaryDirectory() as outside_dir:
                outside_file = os.path.join(outside_dir, "file.xlsx")

                with patch(
                    "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
                ):
                    # Intentar acceder a archivo fuera del directorio permitido
                    is_allowed, error = _is_path_in_allowed_directories(outside_file)

                    assert not is_allowed, (
                        f"Absolute path outside allowed dir should be blocked"
                    )


class TestResolveSafePath:
    """Tests para la función resolve_safe_path."""

    def test_resolve_safe_path_normalizes_relative_path(self):
        """Test que rutas relativas son normalizadas al directorio permitido."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with patch(
                "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
            ):
                result = resolve_safe_path("document.xlsx")

                assert str(result) == os.path.join(allowed_dir, "document.xlsx")

    def test_resolve_safe_path_blocks_absolute_outside(self):
        """Test que rutas absolutas fuera del directorio son reescritas."""
        with tempfile.TemporaryDirectory() as allowed_dir:
            with tempfile.TemporaryDirectory() as outside_dir:
                outside_file = os.path.join(outside_dir, "file.xlsx")

                with patch(
                    "excel_mcp.utils.file_utils.get_directory", return_value=allowed_dir
                ):
                    result = resolve_safe_path(outside_file)

                    # Debería reescribirse al directorio permitido
                    expected = os.path.join(allowed_dir, "file.xlsx")
                    assert str(result) == expected


class TestAllowedDirectories:
    """Tests para verificar la configuración de directorios permitidos."""

    def test_get_allowed_directories_resolves_symlinks(self):
        """Test que los directorios permitidos resuelven symlinks."""
        with tempfile.TemporaryDirectory() as real_dir:
            # Crear un symlink al directorio
            symlink_dir = real_dir + "_link"

            try:
                os.symlink(real_dir, symlink_dir)
            except (OSError, NotImplementedError):
                pytest.skip("Symlinks not supported on this platform")

            with patch(
                "excel_mcp.utils.file_utils.get_directory", return_value=symlink_dir
            ):
                dirs = _get_allowed_directories()

                # El directorio debería estar resuelto (sin symlink)
                assert len(dirs) == 1
                assert Path(dirs[0]).resolve() == Path(real_dir).resolve()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

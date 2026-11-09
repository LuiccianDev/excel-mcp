"""Tests for excel_mcp.utils.cell_utils module."""

import pytest

from excel_mcp.utils.cell_utils import parse_cell_range, validate_cell_reference


class TestParseCellRange:
    """Tests for parse_cell_range function."""

    def test_parse_single_cell_a1(self) -> None:
        """Test parsing single cell reference A1."""
        result = parse_cell_range("A1")
        assert result == (1, 1, None, None)

    def test_parse_single_cell_z100(self) -> None:
        """Test parsing single cell reference Z100."""
        result = parse_cell_range("Z100")
        assert result == (100, 26, None, None)

    def test_parse_range_a1_to_b2(self) -> None:
        """Test parsing range A1:B2."""
        result = parse_cell_range("A1:B2")
        assert result == (1, 1, 2, 2)

    def test_parse_range_aa10_to_bb20(self) -> None:
        """Test parsing range AA10:BB20."""
        result = parse_cell_range("AA10:BB20")
        assert result == (10, 27, 20, 54)

    def test_parse_with_separate_params(self) -> None:
        """Test parsing with start and end as separate parameters."""
        result = parse_cell_range("A1", "D10")
        assert result == (1, 1, 10, 4)

    def test_parse_lowercase_input(self) -> None:
        """Test parsing lowercase cell references."""
        result = parse_cell_range("a1", "c5")
        assert result == (1, 1, 5, 3)

    def test_parse_mixed_case_input(self) -> None:
        """Test parsing mixed case cell references."""
        result = parse_cell_range("Aa1", "Bb2")
        assert result == (1, 27, 2, 54)

    def test_parse_invalid_cell_reference(self) -> None:
        """Test parsing invalid cell reference raises error."""
        with pytest.raises(ValueError) as exc_info:
            parse_cell_range("invalid")

        assert "Invalid cell reference" in str(exc_info.value)

    def test_parse_empty_string(self) -> None:
        """Test parsing empty string raises error."""
        with pytest.raises(ValueError):
            parse_cell_range("")

    def test_parse_no_row_number(self) -> None:
        """Test parsing reference without row number raises error."""
        with pytest.raises(ValueError):
            parse_cell_range("ABC")

    def test_parse_no_column_letter(self) -> None:
        """Test parsing reference without column letter raises error."""
        with pytest.raises(ValueError):
            parse_cell_range("123")

    def test_parse_invalid_end_reference(self) -> None:
        """Test parsing with invalid end reference raises error."""
        with pytest.raises(ValueError) as exc_info:
            parse_cell_range("A1", "invalid")

        assert "Invalid cell reference" in str(exc_info.value)

    def test_parse_large_column_letter(self) -> None:
        """Test parsing large column letter (e.g., XFD)."""
        result = parse_cell_range("XFD1")
        # XFD is the last column in Excel (16384)
        assert result[1] == 16384


class TestValidateCellReference:
    """Tests for validate_cell_reference function."""

    def test_valid_single_cell_a1(self) -> None:
        """Test validating valid cell A1."""
        result = validate_cell_reference("A1")
        assert result is True

    def test_valid_cell_aa100(self) -> None:
        """Test validating valid cell AA100."""
        result = validate_cell_reference("AA100")
        assert result is True

    def test_valid_cell_z999(self) -> None:
        """Test validating valid cell Z999."""
        result = validate_cell_reference("Z999")
        assert result is True

    def test_invalid_empty_string(self) -> None:
        """Test validating empty string."""
        result = validate_cell_reference("")
        assert result is False

    def test_invalid_none(self) -> None:
        """Test validating None."""
        result = validate_cell_reference(None)
        assert result is False

    def test_invalid_letters_after_numbers(self) -> None:
        """Test validating cell with letters after numbers."""
        result = validate_cell_reference("1A")
        assert result is False

    def test_invalid_special_characters(self) -> None:
        """Test validating cell with special characters."""
        invalid_refs = ["A-1", "A_1", "A+1", "@1", "#REF!", "$A$1"]
        for ref in invalid_refs:
            assert validate_cell_reference(ref) is False, (
                f"Expected {ref} to be invalid"
            )

    def test_invalid_only_letters(self) -> None:
        """Test validating reference with only letters."""
        result = validate_cell_reference("ABC")
        assert result is False

    def test_invalid_only_numbers(self) -> None:
        """Test validating reference with only numbers."""
        result = validate_cell_reference("123")
        assert result is False

    def test_invalid_mixed_format(self) -> None:
        """Test validating various invalid formats."""
        invalid_refs = ["A1B2", "12A34", "A 1", "A-1:C2"]
        for ref in invalid_refs:
            assert validate_cell_reference(ref) is False, (
                f"Expected {ref} to be invalid"
            )

    def test_valid_lowercase(self) -> None:
        """Test validating lowercase cell reference."""
        result = validate_cell_reference("a1")
        assert result is True

    def test_valid_mixed_case(self) -> None:
        """Test validating mixed case cell reference."""
        result = validate_cell_reference("Aa1")
        assert result is True

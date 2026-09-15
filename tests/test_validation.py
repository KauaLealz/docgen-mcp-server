import openpyxl
import pytest

from services import bundle_zip_service, parse_range, patch_sheet_service, write_sheet_service


def test_parse_range_rejects_reversed_ranges():
    assert parse_range("B2:A1", 10, 10) is None
    assert parse_range("A0:B2", 10, 10) is None


def test_write_sheet_rejects_non_list_data(tmp_path):
    with pytest.raises(ValueError, match="data deve ser uma lista"):
        write_sheet_service({"path": str(tmp_path / "out.csv"), "data": {"a": 1}})


def test_write_sheet_rejects_invalid_delimiter(tmp_path):
    with pytest.raises(ValueError, match="delimiter"):
        write_sheet_service({"path": str(tmp_path / "out.csv"), "data": [[1]], "delimiter": ";;"})


def test_patch_sheet_rejects_invalid_cell_reference(tmp_path):
    path = tmp_path / "sheet.xlsx"
    openpyxl.Workbook().save(path)
    with pytest.raises(ValueError, match="referência de célula inválida"):
        patch_sheet_service(str(path), [{"cell": "A0", "value": "x"}])


def test_bundle_zip_rejects_duplicate_basenames(tmp_path):
    first_dir = tmp_path / "one"
    second_dir = tmp_path / "two"
    first_dir.mkdir()
    second_dir.mkdir()
    first = first_dir / "same.txt"
    second = second_dir / "same.txt"
    first.write_text("one")
    second.write_text("two")
    with pytest.raises(ValueError, match="mesmo nome"):
        bundle_zip_service([str(first), str(second)], str(tmp_path / "out.zip"))

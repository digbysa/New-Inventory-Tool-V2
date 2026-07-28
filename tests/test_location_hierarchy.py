import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def test_first_location_hierarchy_row_is_not_rejected_as_an_empty_collection():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    function = re.search(
        r"function Add-LocationHierarchyRow \{(?P<body>.*?)\n    \}",
        script,
        re.DOTALL,
    )

    assert function, "Add-LocationHierarchyRow must exist"
    body = function.group("body")
    assert "$null -eq $Rows" in body
    assert "if (-not $Rows" not in body
    assert "$Rows.Add(" in body


def test_location_hierarchy_is_cached_and_uses_authoritative_location_rows():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    function = re.search(
        r"function Get-LocationHierarchyRows \{(?P<body>.*?)\n    \}",
        script,
        re.DOTALL,
    )

    assert function, "Get-LocationHierarchyRows must exist"
    body = function.group("body")
    assert "LocationHierarchyRows" in body
    assert "foreach ($row in @($Inventory.Locations))" in body
    assert "foreach ($row in @($Inventory.Computers))" not in body


def test_department_choices_are_not_restricted_to_the_selected_room():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    assert "Get-UniqueLocationValues -Rows $rows -Property 'Department'" in script
    assert "Get-UniqueLocationValues -Rows $departmentRows -Property 'Department'" not in script


def test_selection_handlers_pass_the_selected_item_not_stale_combo_text():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    for level in ("City", "Location", "Building", "Floor", "Room"):
        expected = f"-ChangedLevel '{level}' -ChangedValue ([string]$ui.{level}ComboBox.SelectedItem)"
        assert expected in script

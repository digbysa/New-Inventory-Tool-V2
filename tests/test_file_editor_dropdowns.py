import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def _file_editor_body():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    function = re.search(
        r"function Show-RoundingEventsFileEditor \{(?P<body>.*?)\n    \}",
        script,
        re.DOTALL,
    )
    assert function, "Show-RoundingEventsFileEditor must exist"
    return function.group("body")


def test_file_editor_uses_yes_no_dropdowns_for_boolean_rounding_columns():
    body = _file_editor_body()

    assert "$yesNoOptions = @('Yes','No')" in body
    for column in (
        "CableMgmnOk",
        "CableMgmtOK",
        "CablingNeeded",
        "LabelOK",
        "CartOK",
        "PeripheralsOK",
        "Rounded",
        "NearbyAdded",
    ):
        assert f"'{column}'" in body
    assert "$yesNoColumns -contains [string]$column" in body
    assert "$comboColumn.ItemsSource = $yesNoOptions" in body


def test_file_editor_always_uses_current_schema_with_nearby_added_last():
    body = _file_editor_body()

    assert "$columns = @($script:RoundingEventColumns)" in body
    assert "$columns = @($imported[0].PSObject.Properties.Name)" not in body


def test_file_editor_uses_requested_maintenance_type_dropdown_options():
    body = _file_editor_body()

    assert (
        "$maintenanceTypeOptions = "
        "@('Mobile Cart','General Rounding','Critical Clinical','Excluded')"
    ) in body
    assert "if ([string]$column -eq 'MaintenanceType')" in body
    assert "$comboColumn.ItemsSource = $maintenanceTypeOptions" in body

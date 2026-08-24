import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def script_text():
    return SCRIPT_PATH.read_text(encoding="utf-8-sig")


def function_body(script, name):
    match = re.search(
        rf"function {re.escape(name)} \{{(?P<body>.*?)\n    \}}",
        script,
        re.DOTALL,
    )
    assert match
    return match.group("body")


def test_inventory_device_keeps_its_original_maintenance_type():
    body = function_body(script_text(), "ConvertTo-DeviceRecord")

    assert "MaintenanceType=Get-FieldValue -Row $Row -Names @('u_device_rounding','MaintenanceType')" in body


def test_selected_device_refreshes_maintenance_type_from_rounding_parent():
    script = script_text()
    updater = function_body(script, "Update-MaintenanceTypeSelection")
    selection = function_body(script, "Set-SelectedSummaryDevice")

    assert "Resolve-ParentDevice -Device $Device -Inventory $Inventory" in updater
    assert "$sourceDevice = if ($parentDevice) { $parentDevice } else { $Device }" in updater
    assert "[string]$sourceDevice.MaintenanceType" in updater
    assert "Update-MaintenanceTypeSelection -Ui $Ui -Device $Device -Inventory $Inventory" in selection


def test_query_does_not_read_removed_raw_csv_property_or_reapply_a_default():
    script = script_text()
    handler = re.search(
        r"\$ui\.QueryButton\.Add_Click\(\{(?P<body>.*?)\n    \}\)",
        script,
        re.DOTALL,
    )

    assert handler
    body = handler.group("body")
    assert "$match.u_device_rounding" not in body
    assert "Set-PrimaryDeviceBindings -Ui $ui -Device $match" in body

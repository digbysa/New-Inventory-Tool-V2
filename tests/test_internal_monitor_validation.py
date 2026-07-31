from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"
SCRIPT = SCRIPT_PATH.read_text(encoding="utf-8-sig")


def test_internal_laptop_panels_are_excluded_from_monitor_validation():
    remote_serials = SCRIPT.split("function Get-RemoteDeviceSerials", 1)[1].split(
        "function Reset-AssociatedSerialValidation", 1
    )[0]

    assert "WmiMonitorConnectionParams" in remote_serials
    assert "$technology -eq 2147483648" in remote_serials
    assert "$internalMonitorInstances.Contains" in remote_serials
    assert "{ continue }" in remote_serials


def test_placeholder_edid_serials_are_not_offered_as_inventory_peripherals():
    remote_serials = SCRIPT.split("function Get-RemoteDeviceSerials", 1)[1].split(
        "function Reset-AssociatedSerialValidation", 1
    )[0]

    assert "$hasTrackableSerial" in remote_serials
    assert "0+|unknown|none|n/?a" in remote_serials
    assert "if ($hasTrackableSerial) { $result.MonitorSerials += $serialText }" in remote_serials
    assert "if ($hasTrackableSerial) { $result.MonitorDetails += $detail }" in remote_serials

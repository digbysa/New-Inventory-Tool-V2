import csv
import re
from pathlib import Path


ROOT = Path(__file__).parents[1]
SCRIPT_PATH = ROOT / "NewAssetTool.Wpf.matching.v14.ps1"


def _function_body(name: str) -> str:
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    match = re.search(
        rf"function {re.escape(name)} \{{(?P<body>.*?)\n    \}}", script, re.DOTALL
    )
    assert match, f"Could not find {name}"
    return match.group("body")


def test_computer_wins_host_name_index_collisions():
    body = _function_body("Add-IndexKey")

    assert "$Value.DetectedType -eq 'Computer'" in body
    assert "$Index[$normalized].DetectedType -ne 'Computer'" in body

    # This is the reported real-world collision: both records are named
    # LD081218, while only the computer row contains summary location data.
    computer_file = ROOT / "Data/Cowichan/Computers - Cowichan.csv"
    monitor_file = ROOT / "Data/Cowichan/Monitors - Cowichan.csv"
    with computer_file.open(encoding="utf-8-sig", newline="") as stream:
        computer = next(row for row in csv.DictReader(stream) if row["name"] == "LD081218")
    with monitor_file.open(encoding="utf-8-sig", newline="") as stream:
        monitor = next(row for row in csv.DictReader(stream) if row["name"] == "LD081218")

    assert monitor["name"] == computer["name"]
    assert computer["u_last_rounded_date"]
    assert computer["location.city"]
    assert computer["u_department_location"]


def test_live_details_reports_progress_and_avoids_strict_mode_count_access():
    lookup = _function_body("Get-LiveComputerDetails")
    click_handler = SCRIPT_PATH.read_text(encoding="utf-8-sig").split(
        "$ui.LiveDetailsButton.Add_Click({", 1
    )[1].split("$ui.MonitorLabelButton.Add_Click({", 1)[0]

    assert ".Count" not in lookup
    assert "Set-StatusMessage -Ui $ui -Mode 'Working'" in click_handler
    assert "DispatcherPriority]::Background" in click_handler
    assert "Set-StatusMessage -Ui $ui -Mode 'Warning'" in click_handler


def test_live_details_applies_timeout_to_cim_session_not_session_options():
    lookup = _function_body("Get-LiveComputerDetails")

    # Windows PowerShell 5.1 does not expose OperationTimeoutSec on
    # New-CimSessionOption. It is a New-CimSession parameter.
    assert "New-CimSessionOption -Protocol Dcom\n" in lookup
    assert (
        "New-CimSession -ComputerName $ComputerName -SessionOption $sessionOption "
        "-OperationTimeoutSec 3 -ErrorAction Stop"
    ) in lookup
    assert "New-CimSessionOption -Protocol Dcom -OperationTimeoutSec" not in lookup


def test_live_details_uses_floating_point_math_for_large_wmi_values():
    body = _function_body("Show-LiveDetailsDialog")

    # Disk and memory sizes commonly exceed Int32.MaxValue. An integer zero as
    # Math.Max's first argument can select the Int32 overload and make the
    # second argument fail conversion before the dialog is displayed.
    assert "[Math]::Max([double]0, [double]$Details.DiskTotal-[double]$Details.DiskFree)" in body
    assert "[Math]::Max([double]0, [double]$Details.RamTotal-[double]$Details.RamFree)" in body


def test_live_details_uses_plain_separators_and_shows_subnet():
    body = _function_body("Show-LiveDetailsDialog")

    assert '"{0}  Build {1}"' in body
    assert "GB used, {1:N0} GB free, {2:N0} GB total" in body
    assert "GB used, {1:N1} GB free, {2:N1} GB total" in body
    assert 'Text="Subnet"' in body
    assert "& $setText 'SubnetValue' $Details.Subnet" in body


def test_live_details_collects_power_battery_and_dock_information():
    lookup = _function_body("Get-LiveComputerDetails")
    dialog = _function_body("Show-LiveDetailsDialog")

    assert "-ClassName Win32_Battery" in lookup
    assert "EstimatedChargeRemaining" in lookup
    assert "'Plugged in'" in lookup
    assert "Batteries: $($charges -join ', ')" in lookup
    assert "-ClassName Win32_PnPEntity" in lookup
    assert "ShowDock=($isLaptop -and $dockDevices.Length -gt 0)" in lookup
    assert 'Text="Power status"' in dialog
    assert 'Text="Docking station"' in dialog
    assert "if ($Details.ShowDock) { 'Visible' } else { 'Collapsed' }" in dialog


def test_live_details_preserves_ipv4_address_and_uses_named_subnet_lookup():
    lookup = _function_body("Get-LiveComputerDetails")

    assert "$addresses = @($network[0].IPAddress)" in lookup
    assert "$ipv4 = [string]$addresses[$i]" in lookup
    assert "Resolve-SubnetName -IpAddress $ipv4 -DataRoot $DataRoot" in lookup
    assert "IPSubnet" not in lookup
    assert "[string]$ipv4[0]" not in lookup


def test_live_details_collects_monitor_and_profile_information():
    lookup = _function_body("Get-LiveComputerDetails")
    dialog = _function_body("Show-LiveDetailsDialog")

    assert "-ClassName WmiMonitorID" in lookup
    assert "-ClassName WmiMonitorConnectionParams" in lookup
    assert "-ClassName WmiMonitorListedSupportedSourceModes" in lookup
    assert '"System DPI: $dpiPercent% ($systemDpi DPI)"' in lookup
    assert "-ClassName Win32_UserProfile" in lookup
    assert 'Text="Monitors"' in dialog
    assert "& $setText 'ProfilesValue' $Details.UserProfileCount" in dialog


def test_live_details_button_tracks_automatic_ping_result():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")

    assert "$Ui.LiveDetailsButton.IsEnabled = $IsOnline" in script
    assert "$ui.LiveDetailsButton.IsEnabled = $false" in script
    assert "LiveDetailsAvailable=$false" in script
    assert "$ui.LiveDetailsButton.IsEnabled = [bool]$script:AppState.LiveDetailsAvailable" in script

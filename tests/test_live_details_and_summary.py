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


def test_live_details_uses_floating_point_math_for_large_wmi_values():
    body = _function_body("Show-LiveDetailsDialog")

    # Disk and memory sizes commonly exceed Int32.MaxValue. An integer zero as
    # Math.Max's first argument can select the Int32 overload and make the
    # second argument fail conversion before the dialog is displayed.
    assert "[Math]::Max([double]0, [double]$Details.DiskTotal-[double]$Details.DiskFree)" in body
    assert "[Math]::Max([double]0, [double]$Details.RamTotal-[double]$Details.RamFree)" in body

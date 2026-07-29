import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def script_text():
    return SCRIPT_PATH.read_text(encoding="utf-8-sig")


def test_saved_system_event_is_marked_rounded_for_nearby_refresh():
    script = script_text()
    save_event = re.search(
        r"function Save-RoundingEvent \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )

    assert save_event
    body = save_event.group("body")
    assert "Rounded='Yes'" in body
    assert "Rounded=$(if ($script:ManualRoundUsed) { 'Yes' } else { 'No' })" not in body


def test_manual_round_is_recorded_only_after_start_process_succeeds():
    script = script_text()
    handler = re.search(
        r"\$ui\.ManualRoundButton\.Add_Click\(\{(?P<body>.*?)\n    \}\)",
        script,
        re.DOTALL,
    )

    assert handler
    body = handler.group("body")
    launch = "Start-Process -FilePath $ui.ManualRoundButton.Tag -ErrorAction Stop"
    success = "$script:ManualRoundUsed = $true"
    assert "$script:ManualRoundUsed = $false" in body
    assert launch in body
    assert success in body
    assert body.index(launch) < body.index(success)
    assert "Unable to open the rounding webpage" in body


def test_save_event_scopes_nearby_from_the_system_location_value():
    script = script_text()
    handler = re.search(
        r"\$ui\.SaveEventButton\.Add_Click\(\{(?P<body>.*?)\n    \}\)",
        script,
        re.DOTALL,
    )

    assert handler
    body = handler.group("body")
    assert "$nearbyScopeDevice = [pscustomobject]@{ Location=[string]$ui.LocationTextBox.Text }" in body
    assert "Add-NearbyScope -Device $nearbyScopeDevice" in body

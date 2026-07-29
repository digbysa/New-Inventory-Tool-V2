import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def script_text():
    return SCRIPT_PATH.read_text(encoding="utf-8-sig")


def test_save_event_defaults_rounded_to_no_unless_manual_page_opened():
    script = script_text()
    save_event = re.search(
        r"function Save-RoundingEvent \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )

    assert save_event
    body = save_event.group("body")
    assert "Rounded=$(if ($script:ManualRoundUsed) { 'Yes' } else { 'No' })" in body
    assert "Rounded='Yes'" not in body


def test_saved_events_are_marked_as_nearby_added_without_changing_rounding_default():
    script = script_text()
    save_event = re.search(
        r"function Save-RoundingEvent \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )
    nearby_save = re.search(
        r"function Save-NearbyEvents \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )

    assert save_event and nearby_save
    assert "NearbyAdded='Yes'" in save_event.group("body")
    assert "Rounded=$(if ($script:ManualRoundUsed) { 'Yes' } else { 'No' })" in save_event.group("body")
    assert "Comments=''; Rounded='No'; NearbyAdded='Yes'" in nearby_save.group("body")


def test_nearby_added_column_is_normalized_to_yes_or_no():
    script = script_text()
    converter = re.search(
        r"function Convert-ToRoundingEventRecord \{(?P<body>.*?)\n    \}",
        script,
        re.DOTALL,
    )

    assert converter
    body = converter.group("body")
    assert "if ($column -eq 'NearbyAdded')" in body
    assert "{ 'Yes' } else { 'No' }" in body


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

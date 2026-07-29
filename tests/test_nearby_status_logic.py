from datetime import datetime, date
from pathlib import Path
import xml.etree.ElementTree as ET

TODAY = date(2026, 7, 24)
CURRENT_ROUNDING_WEEK = {date(2026, 7, 20), date(2026, 7, 21), date(2026, 7, 22), date(2026, 7, 23), TODAY}
XAML_PATH = Path(__file__).resolve().parents[1] / "NewAssetTool.matching.v14.xaml"


def completed(row):
    return str(row.get("NearbyAdded", "")).strip().lower() == "yes"


def most_recent_qualifying(events, asset, maintenance_type):
    key = asset.upper()
    clinical = maintenance_type.strip().lower() == "critical clinical"
    candidates = []
    for event in events:
        if event["AssetTag"].upper() != key or not completed(event):
            continue
        ts = datetime.strptime(event["Timestamp"], "%Y-%m-%d %H:%M:%S")
        if clinical:
            if ts.date() == TODAY:
                candidates.append((ts, event))
        elif ts.date() in CURRENT_ROUNDING_WEEK:
            candidates.append((ts, event))
    return max(candidates, default=(None, None))[1]


def nearby_status(events, asset, maintenance_type):
    event = most_recent_qualifying(events, asset, maintenance_type)
    if event:
        return event.get("CheckStatus") or "Complete", False
    return "-", True


def test_regular_device_uses_most_recent_completed_event_in_current_rounding_week():
    status, editable = nearby_status([
        {"Timestamp": "2026-07-21 08:00:00", "AssetTag": "HSS-1", "CheckStatus": "Inaccessible - Other", "NearbyAdded": "Yes"},
        {"Timestamp": "2026-07-23 09:00:00", "AssetTag": "HSS-1", "CheckStatus": "Inaccessible - In use by Customer", "NearbyAdded": "Yes"},
    ], "hss-1", "General Rounding")
    assert (status, editable) == ("Inaccessible - In use by Customer", False)


def test_regular_device_ignores_previous_week_event_and_keeps_default_editable_status():
    status, editable = nearby_status([
        {"Timestamp": "2026-07-17 09:00:00", "AssetTag": "HSS-2", "CheckStatus": "Inaccessible - Other", "NearbyAdded": "Yes"},
    ], "HSS-2", "General Rounding")
    assert (status, editable) == ("-", True)


def test_clinical_critical_uses_today_event():
    status, editable = nearby_status([
        {"Timestamp": "2026-07-24 10:00:00", "AssetTag": "HSS-3", "CheckStatus": "Inaccessible - Restricted area", "NearbyAdded": "Yes"},
    ], "HSS-3", "Critical Clinical")
    assert (status, editable) == ("Inaccessible - Restricted area", False)


def test_clinical_critical_ignores_previous_day_in_current_week():
    status, editable = nearby_status([
        {"Timestamp": "2026-07-23 10:00:00", "AssetTag": "HSS-4", "CheckStatus": "Inaccessible - Room locked - Key Lock", "NearbyAdded": "Yes"},
    ], "HSS-4", "Critical Clinical")
    assert (status, editable) == ("-", True)


def test_clinical_critical_ignores_previous_week_event():
    status, editable = nearby_status([
        {"Timestamp": "2026-07-16 10:00:00", "AssetTag": "HSS-5", "CheckStatus": "Inaccessible - Other", "NearbyAdded": "Yes"},
    ], "HSS-5", "Critical Clinical")
    assert (status, editable) == ("-", True)


def test_regular_device_uses_latest_nearby_added_event_when_newer_non_nearby_event_exists():
    status, editable = nearby_status([
        {"Timestamp": "2026-07-23 09:00:00", "AssetTag": "HSS-6", "CheckStatus": "Inaccessible - Other", "NearbyAdded": "Yes", "Rounded": "No"},
        {"Timestamp": "2026-07-24 09:00:00", "AssetTag": "HSS-6", "CheckStatus": "Inaccessible - In storage", "NearbyAdded": "No", "Rounded": "Yes"},
    ], "HSS-6", "General Rounding")
    assert (status, editable) == ("Inaccessible - Other", False)


def test_read_only_nearby_status_displays_the_csv_value_as_text():
    root = ET.parse(XAML_PATH).getroot()
    status_column = next(
        element for element in root.iter()
        if element.tag.endswith("DataGridTemplateColumn") and element.get("Header") == "Status"
    )
    status_text = next(
        element for element in status_column.iter()
        if element.tag.endswith("TextBlock") and element.get("Text") == "{Binding Status}"
    )
    trigger = next(
        element for element in status_text.iter()
        if element.tag.endswith("DataTrigger")
        and element.get("Binding") == "{Binding IsStatusEditable}"
        and element.get("Value") == "False"
    )
    assert any(
        element.tag.endswith("Setter")
        and element.get("Property") == "Visibility"
        and element.get("Value") == "Visible"
        for element in trigger
    )

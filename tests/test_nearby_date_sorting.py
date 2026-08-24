from pathlib import Path
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8-sig")
XAML_PATH = ROOT / "NewAssetTool.matching.v14.xaml"


def nearby_column(header):
    root = ET.parse(XAML_PATH).getroot()
    return next(
        element
        for element in root.iter()
        if element.tag.endswith("DataGridTextColumn")
        and element.get("Header") == header
    )


def test_last_rounded_sorts_by_typed_date_instead_of_formatted_text():
    column = nearby_column("Last Rounded")

    assert column.get("Binding") == "{Binding LastRounded}"
    assert column.get("SortMemberPath") == "LastRoundedSort"
    assert "LastRoundedSort=$(if ($lastRoundedDate) { [datetime]$lastRoundedDate }" in SCRIPT


def test_days_ago_sorts_by_integer_instead_of_display_value():
    column = nearby_column("Days Ago")

    assert column.get("Binding") == "{Binding DaysAgo}"
    assert column.get("SortMemberPath") == "DaysAgoSort"
    assert "DaysAgoSort=$(if ($lastRoundedDate) { [int]$daysAgo }" in SCRIPT


def test_rows_without_rounding_dates_have_consistent_typed_sort_sentinels():
    assert "else { [datetime]::MinValue })" in SCRIPT
    assert "else { [int]::MaxValue })" in SCRIPT

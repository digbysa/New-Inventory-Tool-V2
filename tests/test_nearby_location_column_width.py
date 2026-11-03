from pathlib import Path
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8-sig")
XAML_PATH = ROOT / "NewAssetTool.matching.v14.xaml"


def test_nearby_location_column_starts_with_a_flexible_resizable_width():
    root = ET.parse(XAML_PATH).getroot()
    location_column = next(
        element
        for element in root.iter()
        if element.tag.endswith("DataGridTextColumn")
        and element.get("Header") == "Location"
        and element.get("Binding") == "{Binding Location}"
    )

    assert location_column.get("Width") == "*"
    assert location_column.get("CanUserResize") != "False"


def test_nearby_autofit_measures_all_rows_and_defers_hidden_tabs():
    autofit_body = SCRIPT.split("function AutoFit-NearbyColumns", 1)[1].split(
        "function Get-NearbySortState", 1
    )[0]
    assert "-not $grid.IsVisible" in autofit_body
    assert "foreach ($row in $script:NearbyAllRows)" in autofit_body
    assert "$column.Binding.Path.Path" in autofit_body
    assert "$viewer.ViewportWidth" in autofit_body
    assert "DataGridLengthUnitType]::Star" in autofit_body


def test_nearby_status_is_last_and_cells_can_wrap():
    root = ET.parse(XAML_PATH).getroot()
    ns = {"w": "http://schemas.microsoft.com/winfx/2006/xaml/presentation"}
    grid = next(e for e in root.iter() if e.get("{http://schemas.microsoft.com/winfx/2006/xaml}Name") == "NearbyDataGrid")
    columns = list(grid.find("w:DataGrid.Columns", ns))
    assert columns[-1].get("Header") == "Status"
    assert columns[-1].get("Width") == "Auto"
    style = next(e for e in root.iter() if e.get("{http://schemas.microsoft.com/winfx/2006/xaml}Key") == "NearbyCellTextStyle")
    assert any(e.get("Property") == "TextWrapping" and e.get("Value") == "Wrap" for e in style)


def test_nearby_status_reserves_full_option_width_while_location_absorbs_resize():
    body = SCRIPT.split("function AutoFit-NearbyColumns", 1)[1].split("function Get-NearbySortState", 1)[0]
    assert "$measure.Text = 'Inaccessible - Room locked - Card Swipe'" in body
    assert "if ([string]$column.Header -eq 'Status') { $widths[$column] } else { $column.MinWidth }" in body
    layout = body.split("# Share the available width", 1)[1]
    location, status = layout.split("} elseif ([string]$column.Header -eq 'Status')", 1)
    assert "if ([string]$column.Header -eq 'Location')" in location
    assert "DataGridLengthUnitType]::Star" in location
    assert "DataGridLength $widths[$column]" in status


def test_column_measurements_are_cached_and_duplicate_values_are_skipped():
    body = SCRIPT.split("function AutoFit-NearbyColumns", 1)[1].split("function Get-NearbySortState", 1)[0]
    assert "if (-not $script:NearbyNaturalColumnWidths)" in body
    assert "if (-not $seenValues.Add($value)) { continue }" in body
    assert "$script:NearbyNaturalColumnWidths = $widths" in body
    assert "UpdateLayout()" not in body


def test_resize_requests_are_coalesced_without_measuring_in_the_size_handler():
    handler = SCRIPT.split("$ui.NearbyDataGrid.Add_SizeChanged", 1)[1].split("$script:IsUpdatingNearbyFilters", 1)[0]
    assert "Request-NearbyColumnLayout -Ui $ui" in handler
    assert "AutoFit-NearbyColumns" not in handler
    request = SCRIPT.split("function Request-NearbyColumnLayout", 1)[1].split("function AutoFit-NearbyColumns", 1)[0]
    assert "NearbyColumnLayoutTimer.Stop()" in request
    assert "NearbyColumnLayoutTimer.Start()" in request

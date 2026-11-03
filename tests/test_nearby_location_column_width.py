from pathlib import Path
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8-sig")
XAML_PATH = ROOT / "NewAssetTool.matching.v14.xaml"


def test_nearby_location_column_starts_with_a_fixed_resizable_width():
    root = ET.parse(XAML_PATH).getroot()
    location_column = next(
        element
        for element in root.iter()
        if element.tag.endswith("DataGridTextColumn")
        and element.get("Header") == "Location"
        and element.get("Binding") == "{Binding Location}"
    )

    assert location_column.get("Width") == "110"
    assert location_column.get("CanUserResize") != "False"


def test_nearby_autofit_measures_all_rows_and_defers_hidden_tabs():
    autofit_body = SCRIPT.split("function AutoFit-NearbyColumns", 1)[1].split(
        "function Get-NearbySortState", 1
    )[0]
    assert "-not $grid.IsVisible" in autofit_body
    assert "foreach ($row in $grid.ItemsSource)" in autofit_body
    assert "$column.Binding.Path.Path" in autofit_body
    assert "$viewer.ViewportWidth" in autofit_body
    assert "DataGridLengthUnitType]::Star" in autofit_body


def test_nearby_status_fills_last_column_and_cells_can_wrap():
    root = ET.parse(XAML_PATH).getroot()
    ns = {"w": "http://schemas.microsoft.com/winfx/2006/xaml/presentation"}
    grid = next(e for e in root.iter() if e.get("{http://schemas.microsoft.com/winfx/2006/xaml}Name") == "NearbyDataGrid")
    columns = list(grid.find("w:DataGrid.Columns", ns))
    assert columns[-1].get("Header") == "Status"
    assert columns[-1].get("Width") == "*"
    style = next(e for e in root.iter() if e.get("{http://schemas.microsoft.com/winfx/2006/xaml}Key") == "NearbyCellTextStyle")
    assert any(e.get("Property") == "TextWrapping" and e.get("Value") == "Wrap" for e in style)

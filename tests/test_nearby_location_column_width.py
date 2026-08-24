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


def test_nearby_autofit_skips_location_in_measure_and_lock_passes():
    autofit_body = SCRIPT.split("function AutoFit-NearbyColumns", 1)[1].split(
        "function Get-NearbySortState", 1
    )[0]

    assert autofit_body.count("if ([string]$column.Header -eq 'Location') { continue }") == 2
    assert "DataGridLengthUnitType]::SizeToCells" in autofit_body

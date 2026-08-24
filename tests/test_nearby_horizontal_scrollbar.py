from pathlib import Path
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
XAML_PATH = ROOT / "NewAssetTool.matching.v14.xaml"
NS = "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"


def test_nearby_scrollbar_sizes_the_correct_axis_for_each_orientation():
    root = ET.parse(XAML_PATH).getroot()
    nearby_grid = next(
        element
        for element in root.iter(f"{NS}DataGrid")
        if element.get("{http://schemas.microsoft.com/winfx/2006/xaml}Name")
        == "NearbyDataGrid"
    )
    scrollbar_style = next(
        style
        for style in nearby_grid.iter(f"{NS}Style")
        if style.get("TargetType") == "{x:Type ScrollBar}"
    )

    direct_setters = scrollbar_style.findall(f"{NS}Setter")
    assert not any(setter.get("Property") in {"Width", "MinWidth"} for setter in direct_setters)

    triggers = scrollbar_style.findall(f"{NS}Style.Triggers/{NS}Trigger")
    orientation_sizes = {
        trigger.get("Value"): {
            setter.get("Property"): setter.get("Value")
            for setter in trigger.findall(f"{NS}Setter")
        }
        for trigger in triggers
        if trigger.get("Property") == "Orientation"
    }
    assert orientation_sizes["Vertical"] == {"Width": "13", "MinWidth": "13"}
    assert orientation_sizes["Horizontal"] == {"Height": "13", "MinHeight": "13"}
